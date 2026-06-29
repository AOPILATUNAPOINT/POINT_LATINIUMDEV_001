using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	public class frmCierreOperacionesDescuento : System.Windows.Forms.Form
	{
		// --- Controles Autorización ---
		private System.Windows.Forms.GroupBox   grpAutorizacion;
		private System.Windows.Forms.Label      lblNumAutorizacion;
		private System.Windows.Forms.TextBox    txtNumAutorizacion;
		private System.Windows.Forms.Label      lblUsuarioAutoriza;
		private System.Windows.Forms.Label      lblClaveAutoriza;
		private System.Windows.Forms.TextBox    txtClaveAutoriza;
		private System.Windows.Forms.Label      lblFechaAutorizacion;
		private System.Windows.Forms.DateTimePicker dtpFechaAutorizacion;
		private System.Windows.Forms.Button     btnConfirmarAutorizacion;
		private System.Windows.Forms.Label      lblEstadoAutorizacion;

		// --- TabControl ---
		private System.Windows.Forms.TabControl tabControl;
		private System.Windows.Forms.TabPage    tabCargarExcel;
		private System.Windows.Forms.TabPage    tabValidar;
		private System.Windows.Forms.TabPage    tabEjecutar;

		// --- Tab Cargar Excel ---
		private System.Windows.Forms.Label      lblArchivoExcel;
		private System.Windows.Forms.TextBox    txtArchivoExcel;
		private System.Windows.Forms.Button     btnBuscarExcel;
		private System.Windows.Forms.Button     btnCargarExcel;
		private System.Windows.Forms.Button     btnDescargarPlantilla;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridExcel;
		private System.Windows.Forms.Label      lblTotalExcel;

		// --- Tab Validar ---
		private System.Windows.Forms.Button     btnPreparar;
		private System.Windows.Forms.Label      lblContTotal;
		private System.Windows.Forms.Label      lblContEncontrados;
		private System.Windows.Forms.Label      lblContErrores;
		private System.Windows.Forms.TabControl tabValidacion;
		private System.Windows.Forms.TabPage    tabEncontrados;
		private System.Windows.Forms.TabPage    tabErrores;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridEncontrados;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridErrores;

		// --- Tab Ejecutar ---
		private System.Windows.Forms.Button     btnEjecutar;
		private System.Windows.Forms.Label      lblResumenTotal;
		private System.Windows.Forms.Label      lblResumenExitosos;
		private System.Windows.Forms.Label      lblResumenFallidos;
		private System.Windows.Forms.Label      lblResumenMontoProcesado;
		private System.Windows.Forms.Label      lblResumenMontoDescuento;
		private System.Windows.Forms.TabControl tabResultado;
		private System.Windows.Forms.TabPage    tabExitosos;
		private System.Windows.Forms.TabPage    tabFallidos;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridExitosos;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridFallidos;

		// --- Filtros de búsqueda ---
		private System.Windows.Forms.Label      lblFiltroValidos;
		private System.Windows.Forms.TextBox    txtFiltroValidos;
		private System.Windows.Forms.Label      lblFiltroErrores;
		private System.Windows.Forms.TextBox    txtFiltroErrores;
		private System.Windows.Forms.Label      lblFiltroExitosos;
		private System.Windows.Forms.TextBox    txtFiltroExitosos;

		// --- Barra inferior ---
		private System.Windows.Forms.StatusBar     statusBar;
		private System.Windows.Forms.StatusBarPanel sbpMensaje;
		private System.Windows.Forms.ProgressBar   progressBar;

		// --- Seguridad ---
		private Acceso                miAcceso;
		private C1.Data.C1DataSet    cdsSeteoF;

		// --- Estado del flujo ---
		private DataTable dtExcel;
		private bool      autorizacionConfirmada = false;
		private bool      excelCargado           = false;
		private bool      preparacionHecha        = false;
		private int       encontradosCount        = 0;
		private int       erroresCount            = 0;

		private System.ComponentModel.Container components = null;

		public frmCierreOperacionesDescuento()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null) components.Dispose();
			base.Dispose(disposing);
		}

		// ================================================================
		//  INICIALIZACIÓN
		// ================================================================

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCierreOperacionesDescuento_Load(object sender, System.EventArgs e)
		{
			try
			{
				miAcceso  = new Acceso(cdsSeteoF, "09100001");

				if (!Funcion.Permiso("2231", cdsSeteoF))
				{
					MessageBox.Show("No tiene acceso a Cierre de Operaciones con Descuento Especial.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Close();
					return;
				}

				dtpFechaAutorizacion.Value = DateTime.Today;
				SetPlaceholder(txtNumAutorizacion, "Ej: Jimmy Allauca - 08/06/2026");
				lblUsuarioAutoriza.Text = "Usuario: " + MenuLatinium.stUsuario;
				AplicarEstilosGrid();
				ActualizarBotones();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al iniciar el formulario:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private SqlConnection AbrirConexion()
		{
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			SqlConnection cn = new SqlConnection(strConn);
			cn.Open();
			return cn;
		}

		private void SetStatus(string mensaje)
		{
			sbpMensaje.Text = mensaje;
			Application.DoEvents();
		}

		// ================================================================
		//  PLACEHOLDER en TextBox (simulado con Enter/Leave)
		// ================================================================

		private void SetPlaceholder(TextBox txt, string placeholder)
		{
			txt.Text      = placeholder;
			txt.ForeColor = Color.Gray;
			txt.Tag       = placeholder;
		}

		private void Txt_Enter(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (txt.Text == (string)txt.Tag)
			{
				txt.Text      = "";
				txt.ForeColor = Color.Black;
			}
		}

		private void Txt_Leave(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (txt.Text.Trim() == "")
			{
				txt.Text      = (string)txt.Tag;
				txt.ForeColor = Color.Gray;
			}
		}

		private bool TienePlaceholder(TextBox txt)
		{
			return txt.Text == (string)txt.Tag || txt.Text.Trim() == "";
		}

		private void MostrarProgress(bool visible)
		{
			progressBar.Visible = visible;
			if (!visible) progressBar.Value = 0;
		}

		private void ActualizarBotones()
		{
			bool auth = autorizacionConfirmada;

			tabControl.Enabled          = auth;
			btnCargarExcel.Enabled      = auth;
			btnDescargarPlantilla.Enabled = auth;
			btnPreparar.Enabled         = auth && excelCargado;
			btnEjecutar.Enabled         = auth && preparacionHecha && encontradosCount > 0;
		}

		// ================================================================
		//  AUTORIZACIÓN
		// ================================================================

		private void btnConfirmarAutorizacion_Click(object sender, System.EventArgs e)
		{
			if (TienePlaceholder(txtNumAutorizacion))
			{
				MessageBox.Show("Ingrese quién autorizó la operación (nombre y fecha del correo o documento).",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				txtNumAutorizacion.Focus();
				return;
			}
			if (txtClaveAutoriza.Text.Trim() == "")
			{
				MessageBox.Show("Ingrese su clave para confirmar la autorización.",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				txtClaveAutoriza.Focus();
				return;
			}

			// Validar usuario actual con su clave
			SqlConnection cn = null;
			try
			{
				cn = AbrirConexion();
				SqlCommand cmd = new SqlCommand(
					"SELECT Nombre FROM Usuario WHERE Nombre = @Nombre AND Clave = @Clave AND Activo = 1",
					cn);
				cmd.Parameters.Add("@Nombre", SqlDbType.VarChar, 100).Value = MenuLatinium.stUsuario;
				cmd.Parameters.Add("@Clave",  SqlDbType.VarChar, 100).Value = txtClaveAutoriza.Text.Trim();
				object resultado = cmd.ExecuteScalar();

				if (resultado == null || resultado == DBNull.Value)
				{
					MessageBox.Show("Usuario o clave incorrectos. Verifique e intente nuevamente.",
						"Acceso denegado", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					txtClaveAutoriza.Clear();
					txtClaveAutoriza.Focus();
					return;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al validar el usuario:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			finally
			{
				if (cn != null) cn.Close();
			}

			autorizacionConfirmada              = true;
			lblEstadoAutorizacion.Text          = "Autorización confirmada — Usuario: " + MenuLatinium.stUsuario
				+ "  |  Autorizado por: " + txtNumAutorizacion.Text.Trim();
			lblEstadoAutorizacion.ForeColor     = Color.DarkGreen;
			btnConfirmarAutorizacion.Enabled    = false;
			txtNumAutorizacion.Enabled          = false;
			txtClaveAutoriza.Enabled            = false;
			dtpFechaAutorizacion.Enabled        = false;

			ActualizarBotones();
			tabControl.SelectedIndex = 0;
			SetStatus("Autorización confirmada. Puede continuar con el proceso.");
		}

		// ================================================================
		//  TAB 1 — CARGAR EXCEL
		// ================================================================

		private void btnDescargarPlantilla_Click(object sender, System.EventArgs e)
		{
			SaveFileDialog dlg = new SaveFileDialog();
			dlg.Title    = "Guardar plantilla";
			dlg.Filter   = "Excel 97-2003 (*.xls)|*.xls";
			dlg.FileName = "Plantilla_CierreDescuento.xls";
			if (dlg.ShowDialog() != DialogResult.OK) return;
			try
			{
				GenerarPlantilla(dlg.FileName);
				MessageBox.Show("Plantilla guardada en:\n" + dlg.FileName,
					"Plantilla", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al generar plantilla:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void GenerarPlantilla(string rutaArchivo)
		{
			StreamWriter sw = new StreamWriter(rutaArchivo, false, Encoding.UTF8);
			sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sw.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
			sw.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
			sw.WriteLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
			sw.WriteLine(" <Styles>");
			sw.WriteLine("  <Style ss:ID=\"sEnc\"><Font ss:Bold=\"1\"/><Interior ss:Color=\"#003366\" ss:Pattern=\"Solid\"/><Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/></Style>");
			sw.WriteLine("  <Style ss:ID=\"sNota\"><Font ss:Italic=\"1\" ss:Color=\"#CC0000\"/></Style>");
			sw.WriteLine(" </Styles>");
			sw.WriteLine(" <Worksheet ss:Name=\"Descuento\">");
			sw.WriteLine("  <Table>");
			sw.WriteLine("   <Row><Cell ss:Index=\"1\" ss:StyleID=\"sNota\"><Data ss:Type=\"String\">INSTRUCCIONES: No modificar las filas 1 al 2. Datos desde fila 3.</Data></Cell></Row>");
			sw.WriteLine("   <Row>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">ID</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">CEDULA</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">CLIENTE</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">FACTURA</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">SEGMENTO</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEnc\"><Data ss:Type=\"String\">VALOR</Data></Cell>");
			sw.WriteLine("   </Row>");
			for (int i = 1; i <= 10; i++)
			{
				sw.WriteLine("   <Row>");
				sw.WriteLine("    <Cell><Data ss:Type=\"Number\">" + i + "</Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"Number\">0</Data></Cell>");
				sw.WriteLine("   </Row>");
			}
			sw.WriteLine("  </Table>");
			sw.WriteLine(" </Worksheet>");
			sw.WriteLine("</Workbook>");
			sw.Close();
		}

		private void btnBuscarExcel_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog dlg = new OpenFileDialog();
			dlg.Title  = "Seleccionar archivo Excel";
			dlg.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx|CSV (*.csv)|*.csv|Todos|*.*";
			if (dlg.ShowDialog() == DialogResult.OK)
				txtArchivoExcel.Text = dlg.FileName;
		}

		private void btnCargarExcel_Click(object sender, System.EventArgs e)
		{
			if (txtArchivoExcel.Text.Trim() == "")
			{
				MessageBox.Show("Seleccione un archivo Excel.", "Aviso",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			try
			{
				Cursor = Cursors.WaitCursor;
				SetStatus("Leyendo archivo...");
				MostrarProgress(true);

				string archivo = txtArchivoExcel.Text.Trim();
				string ext     = Path.GetExtension(archivo).ToLower();

				if (ext == ".csv")
					LeerCSV(archivo);
				else
					LeerExcelOleDb(archivo);

				gridExcel.DataSource = new DataView(dtExcel);
				AplicarEstiloVisual(gridExcel);
				lblTotalExcel.Text   = "Total leídos: " + dtExcel.Rows.Count;
				excelCargado         = dtExcel.Rows.Count > 0;
				preparacionHecha     = false;

				ActualizarBotones();
				SetStatus("Archivo cargado: " + dtExcel.Rows.Count + " registros.");
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al leer el archivo:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				MostrarProgress(false);
				Cursor = Cursors.Default;
			}
		}

		private void LeerExcelOleDb(string archivo)
		{
			string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + archivo
				+ ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
			OleDbConnection oleConn = new OleDbConnection(connStr);
			try
			{
				oleConn.Open();
				DataTable sheets = oleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
				string hoja = sheets.Rows[0]["TABLE_NAME"].ToString();

				// HDR=No: columnas F1..F6 — fila 1=instrucciones, fila 2=encabezados, datos desde fila 3
				OleDbDataAdapter da = new OleDbDataAdapter(
					"SELECT F1, F2, F3, F4, F5, F6 FROM [" + hoja + "]", oleConn);
				DataTable dtTemp = new DataTable();
				da.Fill(dtTemp);

				dtExcel = ConstruirDataTable();

				for (int i = 2; i < dtTemp.Rows.Count; i++) // desde índice 2 = fila 3
				{
					string id       = dtTemp.Rows[i][0] == DBNull.Value ? "" : dtTemp.Rows[i][0].ToString().Trim();
					string cedula   = dtTemp.Rows[i][1] == DBNull.Value ? "" : dtTemp.Rows[i][1].ToString().Trim();
					string cliente  = dtTemp.Rows[i][2] == DBNull.Value ? "" : dtTemp.Rows[i][2].ToString().Trim();
					string factura  = dtTemp.Rows[i][3] == DBNull.Value ? "" : dtTemp.Rows[i][3].ToString().Trim();
					string segmento = dtTemp.Rows[i][4] == DBNull.Value ? "" : dtTemp.Rows[i][4].ToString().Trim();
					string valorStr = dtTemp.Rows[i][5] == DBNull.Value ? "0" : dtTemp.Rows[i][5].ToString().Trim();

					if (cedula == "" && factura == "") continue;

					decimal valor = 0;
					try { valor = Convert.ToDecimal(valorStr.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture); }
					catch { valor = 0; }

					DataRow nr = dtExcel.NewRow();
					nr["ID"]       = id;
					nr["Cedula"]   = cedula;
					nr["Cliente"]  = cliente;
					nr["Factura"]  = factura;
					nr["Segmento"] = segmento;
					nr["Valor"]    = valor;
					dtExcel.Rows.Add(nr);
				}
			}
			finally
			{
				oleConn.Close();
			}
		}

		private void LeerCSV(string archivo)
		{
			dtExcel = ConstruirDataTable();
			StreamReader sr = new StreamReader(archivo, Encoding.Default);
			int lineaNum = 0;
			string linea;
			while ((linea = sr.ReadLine()) != null)
			{
				lineaNum++;
				if (lineaNum < 3) continue; // fila 1=instrucciones, fila 2=encabezados
				string[] c = linea.Split(',');
				if (c.Length < 6) continue;

				decimal valor = 0;
				try { valor = Convert.ToDecimal(c[5].Trim().Replace("\"", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture); }
				catch { valor = 0; }

				DataRow nr = dtExcel.NewRow();
				nr["ID"]       = c[0].Trim().Replace("\"", "");
				nr["Cedula"]   = c[1].Trim().Replace("\"", "");
				nr["Cliente"]  = c[2].Trim().Replace("\"", "");
				nr["Factura"]  = c[3].Trim().Replace("\"", "");
				nr["Segmento"] = c[4].Trim().Replace("\"", "");
				nr["Valor"]    = valor;
				dtExcel.Rows.Add(nr);
			}
			sr.Close();
		}

		private DataTable ConstruirDataTable()
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("ID",       typeof(string));
			dt.Columns.Add("Cedula",   typeof(string));
			dt.Columns.Add("Cliente",  typeof(string));
			dt.Columns.Add("Factura",  typeof(string));
			dt.Columns.Add("Segmento", typeof(string));
			dt.Columns.Add("Valor",    typeof(decimal));
			return dt;
		}

		// ================================================================
		//  TAB 2 — VALIDAR
		// ================================================================

		private void btnPreparar_Click(object sender, System.EventArgs e)
		{
			if (dtExcel == null || dtExcel.Rows.Count == 0)
			{
				MessageBox.Show("Primero cargue el archivo Excel.", "Aviso",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			SqlConnection cn = null;
			try
			{
				Cursor = Cursors.WaitCursor;
				SetStatus("Validando operaciones...");
				MostrarProgress(true);
				progressBar.Maximum = dtExcel.Rows.Count + 5;
				progressBar.Value   = 0;

				cn = AbrirConexion();

				DataTable dtEnc    = new DataTable("Encontrados");
				dtEnc.Columns.Add("ID",             typeof(string));
				dtEnc.Columns.Add("Cedula",         typeof(string));
				dtEnc.Columns.Add("Cliente",        typeof(string));
				dtEnc.Columns.Add("Factura",        typeof(string));
				dtEnc.Columns.Add("Segmento",       typeof(string));
				dtEnc.Columns.Add("Valor Excel",    typeof(decimal));
				dtEnc.Columns.Add("Saldo Real BD",  typeof(decimal));
				dtEnc.Columns.Add("Monto Condonado",typeof(decimal));

				DataTable dtErr = new DataTable("Errores");
				dtErr.Columns.Add("ID",       typeof(string));
				dtErr.Columns.Add("Cedula",   typeof(string));
				dtErr.Columns.Add("Factura",  typeof(string));
				dtErr.Columns.Add("Segmento", typeof(string));
				dtErr.Columns.Add("Valor",    typeof(decimal));
				dtErr.Columns.Add("Error",    typeof(string));

				foreach (DataRow row in dtExcel.Rows)
				{
					string  id       = row["ID"].ToString();
					string  cedula   = row["Cedula"].ToString().Trim();
					string  cliente  = row["Cliente"].ToString().Trim();
					string  factura  = row["Factura"].ToString().Trim();
					string  segmento = row["Segmento"].ToString().Trim();
					decimal valor    = Convert.ToDecimal(row["Valor"]);

					// Toda la lógica de validación está en el SP —
					// así un cambio de reglas solo requiere actualizar BD, no recompilar.
					SqlCommand cmdVal = new SqlCommand("CierreOperacionDescuentoValidar", cn);
					cmdVal.CommandType    = CommandType.StoredProcedure;
					cmdVal.CommandTimeout = 30;
					cmdVal.Parameters.Add("@Cedula",         SqlDbType.VarChar,  20).Value  = cedula;
					cmdVal.Parameters.Add("@Factura",        SqlDbType.VarChar,  50).Value  = factura;
					cmdVal.Parameters.Add("@Segmento",       SqlDbType.VarChar, 100).Value  = segmento;
					cmdVal.Parameters.Add("@ValorDescuento", SqlDbType.Decimal      ).Value  = valor;

					SqlParameter pIdCompra    = cmdVal.Parameters.Add("@idCompra",      SqlDbType.Int);
					SqlParameter pCliente     = cmdVal.Parameters.Add("@Cliente",        SqlDbType.VarChar, 200);
					SqlParameter pSaldoReal   = cmdVal.Parameters.Add("@SaldoReal",      SqlDbType.Decimal); pSaldoReal.Precision  = 18; pSaldoReal.Scale  = 4;
					SqlParameter pValorAjust  = cmdVal.Parameters.Add("@ValorAjustado",  SqlDbType.Decimal); pValorAjust.Precision = 18; pValorAjust.Scale = 4;
					SqlParameter pError       = cmdVal.Parameters.Add("@Error",           SqlDbType.VarChar, 500);
					pIdCompra  .Direction = ParameterDirection.Output;
					pCliente   .Direction = ParameterDirection.Output;
					pSaldoReal .Direction = ParameterDirection.Output;
					pValorAjust.Direction = ParameterDirection.Output;
					pError     .Direction = ParameterDirection.Output;

					cmdVal.ExecuteNonQuery();

					string  error       = pError     .Value == DBNull.Value ? "" : pError     .Value.ToString();
					decimal saldoReal   = pSaldoReal .Value == DBNull.Value ? 0  : Convert.ToDecimal(pSaldoReal.Value);
					decimal valorFinal  = pValorAjust.Value == DBNull.Value ? valor : Convert.ToDecimal(pValorAjust.Value);
					string  clienteBD   = pCliente   .Value == DBNull.Value ? "" : pCliente   .Value.ToString().Trim();

					if (error == "")
					{
						DataRow nr = dtEnc.NewRow();
						nr["ID"]              = id;
						nr["Cedula"]          = cedula;
						nr["Cliente"]         = clienteBD != "" ? clienteBD : cliente;
						nr["Factura"]         = factura;
						nr["Segmento"]        = segmento;
						nr["Valor Excel"]     = valor;
						nr["Saldo Real BD"]   = saldoReal;
						nr["Monto Condonado"] = 0m;
						dtEnc.Rows.Add(nr);
					}
					else
					{
						DataRow er = dtErr.NewRow();
						er["ID"]       = id;
						er["Cedula"]   = cedula;
						er["Factura"]  = factura;
						er["Segmento"] = segmento;
						er["Valor"]    = valor;
						er["Error"]    = error;
						dtErr.Rows.Add(er);
					}

					progressBar.Value = Math.Min(progressBar.Value + 1, progressBar.Maximum - 1);
					Application.DoEvents();
				}

				gridEncontrados.DataSource = new DataView(dtEnc);
				AplicarEstiloEncontrados(gridEncontrados);
				gridErrores.DataSource = new DataView(dtErr);
				AplicarEstiloErrores(gridErrores);

				encontradosCount = dtEnc.Rows.Count;
				erroresCount     = dtErr.Rows.Count;

				// Calcular totales (Saldo Real BD es el que se procesará)
				decimal totalSaldoBD = 0;
				foreach (DataRow r in dtEnc.Rows)
					totalSaldoBD += Convert.ToDecimal(r["Saldo Real BD"]);

				lblContTotal.Text       = "Total Excel: " + dtExcel.Rows.Count;
				lblContEncontrados.Text = "Válidos: " + encontradosCount
					+ "  |  Saldo Real BD: $" + totalSaldoBD.ToString("N2")
					+ "  (El monto exacto se recalcula al ejecutar)";
				lblContErrores.Text     = "Con errores: "      + erroresCount;
				lblContErrores.ForeColor = erroresCount > 0 ? Color.DarkRed : Color.DarkGreen;

				preparacionHecha = true;
				ActualizarBotones();

				if (erroresCount > 0)
				{
					SetStatus("Validación con errores. Corrija los " + erroresCount + " registros antes de ejecutar.");
					tabValidacion.SelectedIndex = 1;
				}
				else
				{
					SetStatus("Validación exitosa. " + encontradosCount + " operaciones listas para cerrar.");
				}
				progressBar.Value = progressBar.Maximum;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error en validación:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				if (cn != null) cn.Close();
				MostrarProgress(false);
				Cursor = Cursors.Default;
			}
		}

		// Auxiliar para trim en comparación de cédulas
		private string LTRIM(string s) { return s == null ? "" : s.Trim().TrimStart('0'); }

		private void AplicarEstilosGrid() { }

		private void AplicarEstiloVisual(UltraGrid grid)
		{
			Infragistics.Win.Appearance appBack = new Infragistics.Win.Appearance();
			appBack.BackColor = Color.White;
			grid.DisplayLayout.Appearance = appBack;
			grid.DisplayLayout.GroupByBox.Hidden = true;
			grid.DisplayLayout.AutoFitColumns = false;
			grid.DisplayLayout.Override.AllowAddNew = AllowAddNew.No;
			grid.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False;
			grid.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False;
			grid.DisplayLayout.Override.HeaderClickAction = HeaderClickAction.SortSingle;

			Infragistics.Win.Appearance appHeader = new Infragistics.Win.Appearance();
			appHeader.BackColor = Color.FromArgb(89, 135, 214);
			appHeader.BackColor2 = Color.FromArgb(7, 59, 150);
			appHeader.BackGradientStyle = GradientStyle.Vertical;
			appHeader.FontData.BoldAsString = "True";
			appHeader.FontData.Name = "Arial";
			appHeader.FontData.SizeInPoints = 9F;
			appHeader.ForeColor = Color.White;
			appHeader.ThemedElementAlpha = Alpha.Transparent;
			grid.DisplayLayout.Override.HeaderAppearance = appHeader;

			Infragistics.Win.Appearance appAlt = new Infragistics.Win.Appearance();
			appAlt.BackColor = SystemColors.ControlLight;
			appAlt.BackColor2 = Color.Gainsboro;
			appAlt.BackGradientStyle = GradientStyle.Vertical;
			grid.DisplayLayout.Override.RowAlternateAppearance = appAlt;

			Infragistics.Win.Appearance appSel = new Infragistics.Win.Appearance();
			appSel.BackColor = Color.FromArgb(251, 230, 148);
			appSel.BackColor2 = Color.FromArgb(238, 149, 21);
			appSel.BackGradientStyle = GradientStyle.Vertical;
			grid.DisplayLayout.Override.SelectedRowAppearance = appSel;

			grid.FlatMode = true;
			grid.Font = new Font("Tahoma", 7.75F);
		}

		private void AplicarEstiloEncontrados(UltraGrid grid)
		{
			AplicarEstiloVisual(grid);
			if (grid.DisplayLayout.Bands.Count == 0) return;
			UltraGridBand band = grid.DisplayLayout.Bands[0];
			band.Columns["ID"].Width = 40;
			band.Columns["Cedula"].Header.Caption = "Cédula"; band.Columns["Cedula"].Width = 90;
			band.Columns["Cliente"].Width = 160;
			band.Columns["Factura"].Width = 90;
			band.Columns["Segmento"].Width = 120;
			ConfigDecimal(band.Columns["Valor Excel"],    "Valor Excel", 100);
			ConfigDecimal(band.Columns["Saldo Real BD"],  "Condonado",   100);
			ConfigDecimal(band.Columns["Monto Condonado"],"Saldo Final",  90);
		}

		private void AplicarEstiloErrores(UltraGrid grid)
		{
			AplicarEstiloVisual(grid);
			if (grid.DisplayLayout.Bands.Count == 0) return;
			UltraGridBand band = grid.DisplayLayout.Bands[0];
			band.Columns["ID"].Width = 40;
			band.Columns["Cedula"].Header.Caption = "Cédula"; band.Columns["Cedula"].Width = 90;
			band.Columns["Factura"].Width = 90;
			band.Columns["Segmento"].Width = 120;
			ConfigDecimal(band.Columns["Valor"], "Valor", 100);
			band.Columns["Error"].Width = 370;
		}

		private void AplicarEstiloExitosos(UltraGrid grid)
		{
			AplicarEstiloVisual(grid);
			if (grid.DisplayLayout.Bands.Count == 0) return;
			UltraGridBand band = grid.DisplayLayout.Bands[0];
			band.Columns["Factura"].Width = 90;
			band.Columns["Cliente"].Width = 160;
			band.Columns["Segmento"].Width = 120;
			ConfigDecimal(band.Columns["Valor Excel"],   "Valor Excel", 100);
			ConfigDecimal(band.Columns["Saldo Real BD"], "Condonado",   100);
		}

		private void ConfigDecimal(UltraGridColumn col, string caption, int width)
		{
			col.Header.Caption = caption;
			col.Width = width;
			col.Format = "N2";
			col.CellAppearance.TextHAlign = HAlign.Right;
		}

		// ================================================================
		//  FILTROS DE BÚSQUEDA
		// ================================================================

		private void AplicarFiltro(UltraGrid grid, string texto, string columnas)
		{
			if (grid.DataSource == null) return;
			DataView dv = (DataView)grid.DataSource;
			string q = texto.Trim().Replace("'", "''");
			if (q == "")
			{
				dv.RowFilter = "";
				return;
			}
			string[] cols = columnas.Split(',');
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			foreach (string col in cols)
			{
				if (sb.Length > 0) sb.Append(" OR ");
				sb.Append("[" + col.Trim() + "] LIKE '%" + q + "%'");
			}
			dv.RowFilter = sb.ToString();
		}

		private void LimpiarFiltros()
		{
			txtFiltroValidos.Text  = "";
			txtFiltroErrores.Text  = "";
			txtFiltroExitosos.Text = "";
			if (gridEncontrados.DataSource != null) ((DataView)gridEncontrados.DataSource).RowFilter = "";
			if (gridErrores.DataSource     != null) ((DataView)gridErrores.DataSource).RowFilter     = "";
			if (gridExitosos.DataSource    != null) ((DataView)gridExitosos.DataSource).RowFilter    = "";
		}

		private void txtFiltroValidos_TextChanged(object sender, EventArgs e)
		{
			AplicarFiltro(gridEncontrados, txtFiltroValidos.Text, "Cedula,Cliente,Factura,Segmento");
		}

		private void txtFiltroErrores_TextChanged(object sender, EventArgs e)
		{
			AplicarFiltro(gridErrores, txtFiltroErrores.Text, "Cedula,Factura,Segmento,Error");
		}

		private void txtFiltroExitosos_TextChanged(object sender, EventArgs e)
		{
			AplicarFiltro(gridExitosos, txtFiltroExitosos.Text, "Factura,Cliente,Segmento");
		}

		// ================================================================
		//  TAB 3 — EJECUTAR
		// ================================================================

		private void btnEjecutar_Click(object sender, System.EventArgs e)
		{
			if (!preparacionHecha || encontradosCount == 0)
			{
				MessageBox.Show("Complete la validación antes de ejecutar.",
					"Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			// Construir resumen por segmento
			DataView dvResumen = (DataView)gridEncontrados.DataSource;
			dvResumen.RowFilter = "";
			decimal totalMonto = 0;
			System.Collections.Hashtable cntSeg   = new System.Collections.Hashtable();
			System.Collections.Hashtable montoSeg = new System.Collections.Hashtable();
			System.Collections.ArrayList segOrden = new System.Collections.ArrayList();
			foreach (DataRowView r in dvResumen)
			{
				string  seg = r["Segmento"].ToString();
				decimal val = Convert.ToDecimal(r["Saldo Real BD"]);
				totalMonto += val;
				if (!cntSeg.ContainsKey(seg)) { cntSeg[seg] = 0; montoSeg[seg] = 0m; segOrden.Add(seg); }
				cntSeg[seg]   = (int)cntSeg[seg] + 1;
				montoSeg[seg] = (decimal)montoSeg[seg] + val;
			}
			System.Text.StringBuilder sbResumen = new System.Text.StringBuilder();
			sbResumen.Append("==========================================\r\n");
			sbResumen.Append("  RESUMEN DE CIERRE CON DESCUENTO ESPECIAL\r\n");
			sbResumen.Append("==========================================\r\n\r\n");
			sbResumen.Append("  Autorizado por : " + txtNumAutorizacion.Text.Trim() + "\r\n");
			sbResumen.Append("  Usuario        : " + MenuLatinium.stUsuario + "\r\n");
			sbResumen.Append("  Fecha          : " + DateTime.Today.ToString("dd/MM/yyyy") + "\r\n\r\n");
			sbResumen.Append("------------------------------------------\r\n");
			sbResumen.Append("  DETALLE POR SEGMENTO:\r\n");
			sbResumen.Append("------------------------------------------\r\n");
			foreach (string seg in segOrden)
				sbResumen.Append("  " + seg.PadRight(22) + ((int)cntSeg[seg]).ToString().PadLeft(3) + " ops   $" + ((decimal)montoSeg[seg]).ToString("N2").PadLeft(10) + "\r\n");
			sbResumen.Append("------------------------------------------\r\n");
			sbResumen.Append("  TOTAL OPERACIONES : " + encontradosCount + "\r\n");
			sbResumen.Append("  SALDO REAL BD     : $" + totalMonto.ToString("N2") + "\r\n");
			if (erroresCount > 0)
				sbResumen.Append("\r\n  OMITIDOS (con error): " + erroresCount + "\r\n");
			sbResumen.Append("\r\n  NOTA: El monto exacto por operacion\r\n");
			sbResumen.Append("  se recalcula desde la BD al ejecutar.\r\n");
			sbResumen.Append("==========================================\r\n");
			sbResumen.Append("  Esta accion NO se puede revertir.\r\n");
			sbResumen.Append("  Desea continuar?\r\n");
			sbResumen.Append("==========================================");

			if (MessageBox.Show(sbResumen.ToString(),
				"Confirmar Ejecución",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
				return;

			DataTable dtExitosos = new DataTable("Exitosos");
			dtExitosos.Columns.Add("Factura",       typeof(string));
			dtExitosos.Columns.Add("Cliente",       typeof(string));
			dtExitosos.Columns.Add("Segmento",      typeof(string));
			dtExitosos.Columns.Add("Valor Excel",   typeof(decimal));
			dtExitosos.Columns.Add("Saldo Real BD", typeof(decimal));

			DataTable dtFallidos = new DataTable();
			dtFallidos.Columns.Add("Factura",  typeof(string));
			dtFallidos.Columns.Add("Cedula",   typeof(string));
			dtFallidos.Columns.Add("Segmento", typeof(string));
			dtFallidos.Columns.Add("Error",    typeof(string));

			// Limpiar filtros para procesar TODOS los registros válidos
			LimpiarFiltros();
			DataView dvEnc = (DataView)gridEncontrados.DataSource;
			dvEnc.RowFilter = "";
			int total      = dvEnc.Count;
			int exitosos   = 0;
			int fallidos   = 0;

			SqlConnection cn = null;
			try
			{
				Cursor = Cursors.WaitCursor;
				SetStatus("Ejecutando cierre de operaciones...");
				MostrarProgress(true);
				progressBar.Maximum = total;
				progressBar.Value   = 0;

				cn = AbrirConexion();

				foreach (DataRowView drv in dvEnc)
				{
					string  factura  = drv["Factura"].ToString();
					string  cedula   = drv["Cedula"].ToString();
					string  cliente  = drv["Cliente"].ToString();
					string  segmento = drv["Segmento"].ToString();
					decimal valor    = Convert.ToDecimal(drv["Saldo Real BD"]);
					decimal valorExcel = Convert.ToDecimal(drv["Valor Excel"]);

					string mensaje = "";
					try
					{
						SqlCommand cmd = new SqlCommand("CierreOperacionDescuento", cn);
						cmd.CommandType    = CommandType.StoredProcedure;
						cmd.CommandTimeout = 60;
						cmd.Parameters.Add("@Cedula",    SqlDbType.VarChar,  20).Value = cedula;
						cmd.Parameters.Add("@Factura",   SqlDbType.VarChar,  50).Value = factura;
						cmd.Parameters.Add("@Segmento",  SqlDbType.VarChar, 100).Value = segmento;
						SqlParameter pValDesc = cmd.Parameters.Add("@ValorDescuento", SqlDbType.Decimal); pValDesc.Precision = 18; pValDesc.Scale = 4; pValDesc.Value = valor;
						cmd.Parameters.Add("@NumAutorizacion", SqlDbType.VarChar, 100).Value = txtNumAutorizacion.Text.Trim();
						cmd.Parameters.Add("@UsuarioAutoriza", SqlDbType.VarChar, 50).Value  = MenuLatinium.stUsuario;
						cmd.Parameters.Add("@UsuarioEjecuta",  SqlDbType.VarChar, 50).Value  = MenuLatinium.stUsuario;
						SqlParameter pMsg = cmd.Parameters.Add("@Mensaje", SqlDbType.VarChar, 500);
						pMsg.Direction = ParameterDirection.Output;
						cmd.ExecuteNonQuery();
						mensaje = pMsg.Value == DBNull.Value ? "" : pMsg.Value.ToString();
					}
					catch (Exception exCmd)
					{
						mensaje = "ERROR: " + exCmd.Message;
					}

					if (mensaje == "")
					{
						DataRow nr = dtExitosos.NewRow();
						nr["Factura"]        = factura;
						nr["Cliente"]        = cliente;
						nr["Segmento"]       = segmento;
						nr["Valor Excel"]   = valorExcel;
						nr["Saldo Real BD"] = valor;
						dtExitosos.Rows.Add(nr);
						exitosos++;
					}
					else
					{
						DataRow nr = dtFallidos.NewRow();
						nr["Factura"]  = factura;
						nr["Cedula"]   = cedula;
						nr["Segmento"] = segmento;
						nr["Error"]    = mensaje;
						dtFallidos.Rows.Add(nr);
						fallidos++;
					}

					progressBar.Value = Math.Min(progressBar.Value + 1, progressBar.Maximum);
					Application.DoEvents();
				}

				gridExitosos.DataSource = new DataView(dtExitosos);
				AplicarEstiloExitosos(gridExitosos);
				gridFallidos.DataSource = new DataView(dtFallidos);
				AplicarEstiloVisual(gridFallidos);

				decimal mProcesado = 0, mExcel = 0;
				foreach (DataRow r in dtExitosos.Rows)
				{
					mProcesado += Convert.ToDecimal(r["Saldo Real BD"]);
					mExcel     += Convert.ToDecimal(r["Valor Excel"]);
				}

				lblResumenTotal.Text          = "Total procesadas: "    + total;
				lblResumenExitosos.Text       = "Exitosas: "            + exitosos;
				lblResumenFallidos.Text       = "Fallidas: "            + fallidos;
				lblResumenMontoProcesado.Text = "Saldo Real BD: $"      + mProcesado.ToString("N2");
				lblResumenMontoDescuento.Text = "Valor Excel referencial: $" + mExcel.ToString("N2");
				lblResumenFallidos.ForeColor  = fallidos > 0 ? Color.DarkRed : Color.DarkGreen;

				tabControl.SelectedIndex = 2;
				if (fallidos > 0)
					tabResultado.SelectedIndex = 1;

				SetStatus("Proceso terminado. Exitosas: " + exitosos + "  |  Fallidas: " + fallidos);

				string resumenFinal =
					"CIERRE DE OPERACIONES CON DESCUENTO ESPECIAL\r\n" +
					"============================================\r\n\r\n" +
					"  Fecha de ejecución : " + DateTime.Now.ToString("dd/MM/yyyy  HH:mm:ss") + "\r\n" +
					"  Ejecutado por      : " + MenuLatinium.stUsuario + "\r\n" +
					"  Autorizado por     : " + txtNumAutorizacion.Text.Trim() + "\r\n\r\n" +
					"--------------------------------------------\r\n" +
					"  QUÉ SE REALIZÓ:\r\n" +
					"  Para cada operación exitosa se:\r\n" +
					"   • Creó un anticipo de pago (Bodega 57)\r\n" +
					"   • Aplicó el saldo real sobre las cuotas\r\n" +
					"   • Condonó las cuotas restantes\r\n" +
					"   • Generó el asiento contable automático\r\n" +
					"   • Registró en el log de auditoría\r\n\r\n" +
					"--------------------------------------------\r\n" +
					"  RESULTADO:\r\n" +
					"  Total procesadas   : " + total + "\r\n" +
					"  Exitosas           : " + exitosos + "\r\n" +
					"  Fallidas           : " + fallidos + "\r\n\r\n" +
					"  Saldo Real BD      : $" + mProcesado.ToString("N2") + "\r\n" +
					"  Valor Excel ref.   : $" + mExcel.ToString("N2") + "\r\n";

				if (fallidos > 0)
					resumenFinal +=
						"\r\n--------------------------------------------\r\n" +
						"  ATENCIÓN: " + fallidos + " operación(es) NO se procesaron.\r\n" +
						"  Revise la pestaña \"Fallidas\" para ver\r\n" +
						"  el detalle de cada error.\r\n";

				resumenFinal += "============================================";

				MessageBox.Show(resumenFinal, "Resultado Final", MessageBoxButtons.OK,
					fallidos > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error en ejecución:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				if (cn != null) cn.Close();
				MostrarProgress(false);
				Cursor = Cursors.Default;
			}
		}

		// ================================================================
		//  INITIALIZE COMPONENT
		// ================================================================

		private void InitializeComponent()
		{
			this.grpAutorizacion = new System.Windows.Forms.GroupBox();
			this.lblNumAutorizacion = new System.Windows.Forms.Label();
			this.txtNumAutorizacion = new System.Windows.Forms.TextBox();
			this.lblUsuarioAutoriza = new System.Windows.Forms.Label();
			this.lblClaveAutoriza = new System.Windows.Forms.Label();
			this.txtClaveAutoriza = new System.Windows.Forms.TextBox();
			this.lblFechaAutorizacion = new System.Windows.Forms.Label();
			this.dtpFechaAutorizacion = new System.Windows.Forms.DateTimePicker();
			this.btnConfirmarAutorizacion = new System.Windows.Forms.Button();
			this.lblEstadoAutorizacion = new System.Windows.Forms.Label();
			this.tabControl = new System.Windows.Forms.TabControl();
			this.tabCargarExcel = new System.Windows.Forms.TabPage();
			this.lblArchivoExcel = new System.Windows.Forms.Label();
			this.txtArchivoExcel = new System.Windows.Forms.TextBox();
			this.btnBuscarExcel = new System.Windows.Forms.Button();
			this.btnCargarExcel = new System.Windows.Forms.Button();
			this.btnDescargarPlantilla = new System.Windows.Forms.Button();
			this.gridExcel = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblTotalExcel = new System.Windows.Forms.Label();
			this.tabValidar = new System.Windows.Forms.TabPage();
			this.btnPreparar = new System.Windows.Forms.Button();
			this.lblContTotal = new System.Windows.Forms.Label();
			this.lblContEncontrados = new System.Windows.Forms.Label();
			this.lblContErrores = new System.Windows.Forms.Label();
			this.tabValidacion = new System.Windows.Forms.TabControl();
			this.tabEncontrados = new System.Windows.Forms.TabPage();
			this.lblFiltroValidos = new System.Windows.Forms.Label();
			this.txtFiltroValidos = new System.Windows.Forms.TextBox();
			this.gridEncontrados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.tabErrores = new System.Windows.Forms.TabPage();
			this.lblFiltroErrores = new System.Windows.Forms.Label();
			this.txtFiltroErrores = new System.Windows.Forms.TextBox();
			this.gridErrores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.tabEjecutar = new System.Windows.Forms.TabPage();
			this.btnEjecutar = new System.Windows.Forms.Button();
			this.lblResumenTotal = new System.Windows.Forms.Label();
			this.lblResumenExitosos = new System.Windows.Forms.Label();
			this.lblResumenFallidos = new System.Windows.Forms.Label();
			this.lblResumenMontoProcesado = new System.Windows.Forms.Label();
			this.lblResumenMontoDescuento = new System.Windows.Forms.Label();
			this.tabResultado = new System.Windows.Forms.TabControl();
			this.tabExitosos = new System.Windows.Forms.TabPage();
			this.lblFiltroExitosos = new System.Windows.Forms.Label();
			this.txtFiltroExitosos = new System.Windows.Forms.TextBox();
			this.gridExitosos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.tabFallidos = new System.Windows.Forms.TabPage();
			this.gridFallidos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.statusBar = new System.Windows.Forms.StatusBar();
			this.sbpMensaje = new System.Windows.Forms.StatusBarPanel();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.grpAutorizacion.SuspendLayout();
			this.tabControl.SuspendLayout();
			this.tabCargarExcel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridExcel)).BeginInit();
			this.tabValidar.SuspendLayout();
			this.tabValidacion.SuspendLayout();
			this.tabEncontrados.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridEncontrados)).BeginInit();
			this.tabErrores.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridErrores)).BeginInit();
			this.tabEjecutar.SuspendLayout();
			this.tabResultado.SuspendLayout();
			this.tabExitosos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridExitosos)).BeginInit();
			this.tabFallidos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridFallidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grpAutorizacion
			// 
			this.grpAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grpAutorizacion.Controls.Add(this.lblNumAutorizacion);
			this.grpAutorizacion.Controls.Add(this.txtNumAutorizacion);
			this.grpAutorizacion.Controls.Add(this.lblUsuarioAutoriza);
			this.grpAutorizacion.Controls.Add(this.lblClaveAutoriza);
			this.grpAutorizacion.Controls.Add(this.txtClaveAutoriza);
			this.grpAutorizacion.Controls.Add(this.lblFechaAutorizacion);
			this.grpAutorizacion.Controls.Add(this.dtpFechaAutorizacion);
			this.grpAutorizacion.Controls.Add(this.btnConfirmarAutorizacion);
			this.grpAutorizacion.Controls.Add(this.lblEstadoAutorizacion);
			this.grpAutorizacion.Location = new System.Drawing.Point(8, 8);
			this.grpAutorizacion.Name = "grpAutorizacion";
			this.grpAutorizacion.Size = new System.Drawing.Size(870, 80);
			this.grpAutorizacion.TabIndex = 0;
			this.grpAutorizacion.TabStop = false;
			this.grpAutorizacion.Text = "Autorización Financiera  (requerida antes de continuar)";
			// 
			// lblNumAutorizacion
			// 
			this.lblNumAutorizacion.Location = new System.Drawing.Point(10, 22);
			this.lblNumAutorizacion.Name = "lblNumAutorizacion";
			this.lblNumAutorizacion.Size = new System.Drawing.Size(95, 18);
			this.lblNumAutorizacion.TabIndex = 0;
			this.lblNumAutorizacion.Text = "Autorizado por:";
			// 
			// txtNumAutorizacion
			// 
			this.txtNumAutorizacion.Location = new System.Drawing.Point(107, 20);
			this.txtNumAutorizacion.Name = "txtNumAutorizacion";
			this.txtNumAutorizacion.Size = new System.Drawing.Size(220, 20);
			this.txtNumAutorizacion.TabIndex = 1;
			this.txtNumAutorizacion.Text = "";
			this.txtNumAutorizacion.Leave += new System.EventHandler(this.Txt_Leave);
			this.txtNumAutorizacion.Enter += new System.EventHandler(this.Txt_Enter);
			// 
			// lblUsuarioAutoriza
			// 
			this.lblUsuarioAutoriza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblUsuarioAutoriza.Location = new System.Drawing.Point(337, 22);
			this.lblUsuarioAutoriza.Name = "lblUsuarioAutoriza";
			this.lblUsuarioAutoriza.Size = new System.Drawing.Size(200, 18);
			this.lblUsuarioAutoriza.TabIndex = 2;
			this.lblUsuarioAutoriza.Text = "Usuario:";
			// 
			// lblClaveAutoriza
			// 
			this.lblClaveAutoriza.Location = new System.Drawing.Point(545, 22);
			this.lblClaveAutoriza.Name = "lblClaveAutoriza";
			this.lblClaveAutoriza.Size = new System.Drawing.Size(40, 18);
			this.lblClaveAutoriza.TabIndex = 3;
			this.lblClaveAutoriza.Text = "Clave:";
			// 
			// txtClaveAutoriza
			// 
			this.txtClaveAutoriza.Location = new System.Drawing.Point(587, 20);
			this.txtClaveAutoriza.Name = "txtClaveAutoriza";
			this.txtClaveAutoriza.PasswordChar = '*';
			this.txtClaveAutoriza.Size = new System.Drawing.Size(110, 20);
			this.txtClaveAutoriza.TabIndex = 4;
			this.txtClaveAutoriza.Text = "";
			// 
			// lblFechaAutorizacion
			// 
			this.lblFechaAutorizacion.Location = new System.Drawing.Point(715, 22);
			this.lblFechaAutorizacion.Name = "lblFechaAutorizacion";
			this.lblFechaAutorizacion.Size = new System.Drawing.Size(40, 18);
			this.lblFechaAutorizacion.TabIndex = 5;
			this.lblFechaAutorizacion.Text = "Fecha:";
			// 
			// dtpFechaAutorizacion
			// 
			this.dtpFechaAutorizacion.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpFechaAutorizacion.Location = new System.Drawing.Point(757, 20);
			this.dtpFechaAutorizacion.Name = "dtpFechaAutorizacion";
			this.dtpFechaAutorizacion.Size = new System.Drawing.Size(90, 20);
			this.dtpFechaAutorizacion.TabIndex = 6;
			// 
			// btnConfirmarAutorizacion
			// 
			this.btnConfirmarAutorizacion.BackColor = System.Drawing.Color.DarkGreen;
			this.btnConfirmarAutorizacion.ForeColor = System.Drawing.Color.White;
			this.btnConfirmarAutorizacion.Location = new System.Drawing.Point(768, 44);
			this.btnConfirmarAutorizacion.Name = "btnConfirmarAutorizacion";
			this.btnConfirmarAutorizacion.Size = new System.Drawing.Size(88, 23);
			this.btnConfirmarAutorizacion.TabIndex = 7;
			this.btnConfirmarAutorizacion.Text = "Confirmar";
			this.btnConfirmarAutorizacion.Click += new System.EventHandler(this.btnConfirmarAutorizacion_Click);
			// 
			// lblEstadoAutorizacion
			// 
			this.lblEstadoAutorizacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblEstadoAutorizacion.ForeColor = System.Drawing.Color.Gray;
			this.lblEstadoAutorizacion.Location = new System.Drawing.Point(10, 52);
			this.lblEstadoAutorizacion.Name = "lblEstadoAutorizacion";
			this.lblEstadoAutorizacion.Size = new System.Drawing.Size(780, 18);
			this.lblEstadoAutorizacion.TabIndex = 8;
			this.lblEstadoAutorizacion.Text = "Pendiente de confirmación";
			// 
			// tabControl
			// 
			this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl.Controls.Add(this.tabCargarExcel);
			this.tabControl.Controls.Add(this.tabValidar);
			this.tabControl.Controls.Add(this.tabEjecutar);
			this.tabControl.Enabled = false;
			this.tabControl.Location = new System.Drawing.Point(8, 96);
			this.tabControl.Name = "tabControl";
			this.tabControl.SelectedIndex = 0;
			this.tabControl.Size = new System.Drawing.Size(870, 520);
			this.tabControl.TabIndex = 1;
			// 
			// tabCargarExcel
			// 
			this.tabCargarExcel.Controls.Add(this.lblArchivoExcel);
			this.tabCargarExcel.Controls.Add(this.txtArchivoExcel);
			this.tabCargarExcel.Controls.Add(this.btnBuscarExcel);
			this.tabCargarExcel.Controls.Add(this.btnCargarExcel);
			this.tabCargarExcel.Controls.Add(this.btnDescargarPlantilla);
			this.tabCargarExcel.Controls.Add(this.gridExcel);
			this.tabCargarExcel.Controls.Add(this.lblTotalExcel);
			this.tabCargarExcel.Location = new System.Drawing.Point(4, 22);
			this.tabCargarExcel.Name = "tabCargarExcel";
			this.tabCargarExcel.Size = new System.Drawing.Size(862, 494);
			this.tabCargarExcel.TabIndex = 0;
			this.tabCargarExcel.Text = "1. Cargar Excel";
			// 
			// lblArchivoExcel
			// 
			this.lblArchivoExcel.Location = new System.Drawing.Point(8, 14);
			this.lblArchivoExcel.Name = "lblArchivoExcel";
			this.lblArchivoExcel.Size = new System.Drawing.Size(55, 18);
			this.lblArchivoExcel.TabIndex = 0;
			this.lblArchivoExcel.Text = "Archivo:";
			// 
			// txtArchivoExcel
			// 
			this.txtArchivoExcel.Location = new System.Drawing.Point(65, 12);
			this.txtArchivoExcel.Name = "txtArchivoExcel";
			this.txtArchivoExcel.ReadOnly = true;
			this.txtArchivoExcel.Size = new System.Drawing.Size(560, 20);
			this.txtArchivoExcel.TabIndex = 1;
			this.txtArchivoExcel.Text = "";
			// 
			// btnBuscarExcel
			// 
			this.btnBuscarExcel.Location = new System.Drawing.Point(632, 11);
			this.btnBuscarExcel.Name = "btnBuscarExcel";
			this.btnBuscarExcel.Size = new System.Drawing.Size(30, 23);
			this.btnBuscarExcel.TabIndex = 2;
			this.btnBuscarExcel.Text = "...";
			this.btnBuscarExcel.Click += new System.EventHandler(this.btnBuscarExcel_Click);
			// 
			// btnCargarExcel
			// 
			this.btnCargarExcel.Location = new System.Drawing.Point(670, 11);
			this.btnCargarExcel.Name = "btnCargarExcel";
			this.btnCargarExcel.Size = new System.Drawing.Size(70, 23);
			this.btnCargarExcel.TabIndex = 3;
			this.btnCargarExcel.Text = "Cargar";
			this.btnCargarExcel.Click += new System.EventHandler(this.btnCargarExcel_Click);
			// 
			// btnDescargarPlantilla
			// 
			this.btnDescargarPlantilla.Location = new System.Drawing.Point(748, 11);
			this.btnDescargarPlantilla.Name = "btnDescargarPlantilla";
			this.btnDescargarPlantilla.Size = new System.Drawing.Size(100, 23);
			this.btnDescargarPlantilla.TabIndex = 4;
			this.btnDescargarPlantilla.Text = "Plantilla";
			this.btnDescargarPlantilla.Click += new System.EventHandler(this.btnDescargarPlantilla_Click);
			// 
			// gridExcel
			// 
			this.gridExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridExcel.Location = new System.Drawing.Point(8, 42);
			this.gridExcel.Name = "gridExcel";
			this.gridExcel.Size = new System.Drawing.Size(846, 440);
			this.gridExcel.TabIndex = 5;
			// 
			// lblTotalExcel
			// 
			this.lblTotalExcel.Location = new System.Drawing.Point(8, 468);
			this.lblTotalExcel.Name = "lblTotalExcel";
			this.lblTotalExcel.Size = new System.Drawing.Size(300, 18);
			this.lblTotalExcel.TabIndex = 6;
			this.lblTotalExcel.Text = "Total leídos: 0";
			// 
			// tabValidar
			// 
			this.tabValidar.Controls.Add(this.btnPreparar);
			this.tabValidar.Controls.Add(this.lblContTotal);
			this.tabValidar.Controls.Add(this.lblContEncontrados);
			this.tabValidar.Controls.Add(this.lblContErrores);
			this.tabValidar.Controls.Add(this.tabValidacion);
			this.tabValidar.Location = new System.Drawing.Point(4, 22);
			this.tabValidar.Name = "tabValidar";
			this.tabValidar.Size = new System.Drawing.Size(862, 494);
			this.tabValidar.TabIndex = 1;
			this.tabValidar.Text = "2. Validar";
			// 
			// btnPreparar
			// 
			this.btnPreparar.BackColor = System.Drawing.Color.SteelBlue;
			this.btnPreparar.ForeColor = System.Drawing.Color.White;
			this.btnPreparar.Location = new System.Drawing.Point(8, 8);
			this.btnPreparar.Name = "btnPreparar";
			this.btnPreparar.Size = new System.Drawing.Size(150, 28);
			this.btnPreparar.TabIndex = 0;
			this.btnPreparar.Text = "Preparar y Validar";
			this.btnPreparar.Click += new System.EventHandler(this.btnPreparar_Click);
			// 
			// lblContTotal
			// 
			this.lblContTotal.Location = new System.Drawing.Point(170, 14);
			this.lblContTotal.Name = "lblContTotal";
			this.lblContTotal.Size = new System.Drawing.Size(150, 18);
			this.lblContTotal.TabIndex = 1;
			this.lblContTotal.Text = "Total Excel: 0";
			// 
			// lblContEncontrados
			// 
			this.lblContEncontrados.Location = new System.Drawing.Point(330, 14);
			this.lblContEncontrados.Name = "lblContEncontrados";
			this.lblContEncontrados.Size = new System.Drawing.Size(360, 18);
			this.lblContEncontrados.TabIndex = 2;
			this.lblContEncontrados.Text = "Válidos: 0";
			// 
			// lblContErrores
			// 
			this.lblContErrores.Location = new System.Drawing.Point(700, 14);
			this.lblContErrores.Name = "lblContErrores";
			this.lblContErrores.Size = new System.Drawing.Size(140, 18);
			this.lblContErrores.TabIndex = 3;
			this.lblContErrores.Text = "Con errores: 0";
			//
			// tabValidacion
			//
			this.tabValidacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
				| System.Windows.Forms.AnchorStyles.Left)
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabValidacion.Controls.Add(this.tabEncontrados);
			this.tabValidacion.Controls.Add(this.tabErrores);
			this.tabValidacion.Location = new System.Drawing.Point(8, 44);
			this.tabValidacion.Name = "tabValidacion";
			this.tabValidacion.SelectedIndex = 0;
			this.tabValidacion.Size = new System.Drawing.Size(846, 442);
			this.tabValidacion.TabIndex = 4;
			// 
			// tabEncontrados
			//
			this.tabEncontrados.Controls.Add(this.lblFiltroValidos);
			this.tabEncontrados.Controls.Add(this.txtFiltroValidos);
			this.tabEncontrados.Controls.Add(this.gridEncontrados);
			this.tabEncontrados.Location = new System.Drawing.Point(0, 0);
			this.tabEncontrados.Name = "tabEncontrados";
			this.tabEncontrados.Size = new System.Drawing.Size(838, 420);
			this.tabEncontrados.TabIndex = 0;
			this.tabEncontrados.Text = "Válidos";
			// 
			// lblFiltroValidos
			// 
			this.lblFiltroValidos.Location = new System.Drawing.Point(2, 4);
			this.lblFiltroValidos.Name = "lblFiltroValidos";
			this.lblFiltroValidos.Size = new System.Drawing.Size(45, 18);
			this.lblFiltroValidos.TabIndex = 0;
			this.lblFiltroValidos.Text = "Buscar:";
			// 
			// txtFiltroValidos
			// 
			this.txtFiltroValidos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFiltroValidos.Location = new System.Drawing.Point(50, 2);
			this.txtFiltroValidos.Name = "txtFiltroValidos";
			this.txtFiltroValidos.Size = new System.Drawing.Size(760, 20);
			this.txtFiltroValidos.TabIndex = 1;
			this.txtFiltroValidos.Text = "";
			this.txtFiltroValidos.TextChanged += new System.EventHandler(this.txtFiltroValidos_TextChanged);
			// 
			// gridEncontrados
			// 
			this.gridEncontrados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridEncontrados.Location = new System.Drawing.Point(0, 28);
			this.gridEncontrados.Name = "gridEncontrados";
			this.gridEncontrados.Size = new System.Drawing.Size(838, 392);
			this.gridEncontrados.TabIndex = 2;
			// 
			// tabErrores
			// 
			this.tabErrores.Controls.Add(this.lblFiltroErrores);
			this.tabErrores.Controls.Add(this.txtFiltroErrores);
			this.tabErrores.Controls.Add(this.gridErrores);
			this.tabErrores.Location = new System.Drawing.Point(0, 0);
			this.tabErrores.Name = "tabErrores";
			this.tabErrores.Size = new System.Drawing.Size(838, 420);
			this.tabErrores.TabIndex = 1;
			this.tabErrores.Text = "Con Errores";
			// 
			// lblFiltroErrores
			// 
			this.lblFiltroErrores.Location = new System.Drawing.Point(2, 4);
			this.lblFiltroErrores.Name = "lblFiltroErrores";
			this.lblFiltroErrores.Size = new System.Drawing.Size(45, 18);
			this.lblFiltroErrores.TabIndex = 0;
			this.lblFiltroErrores.Text = "Buscar:";
			// 
			// txtFiltroErrores
			// 
			this.txtFiltroErrores.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFiltroErrores.Location = new System.Drawing.Point(50, 2);
			this.txtFiltroErrores.Name = "txtFiltroErrores";
			this.txtFiltroErrores.Size = new System.Drawing.Size(760, 20);
			this.txtFiltroErrores.TabIndex = 1;
			this.txtFiltroErrores.Text = "";
			this.txtFiltroErrores.TextChanged += new System.EventHandler(this.txtFiltroErrores_TextChanged);
			// 
			// gridErrores
			// 
			this.gridErrores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridErrores.Location = new System.Drawing.Point(0, 28);
			this.gridErrores.Name = "gridErrores";
			this.gridErrores.Size = new System.Drawing.Size(838, 392);
			this.gridErrores.TabIndex = 2;
			// 
			// tabEjecutar
			// 
			this.tabEjecutar.Controls.Add(this.btnEjecutar);
			this.tabEjecutar.Controls.Add(this.lblResumenTotal);
			this.tabEjecutar.Controls.Add(this.lblResumenExitosos);
			this.tabEjecutar.Controls.Add(this.lblResumenFallidos);
			this.tabEjecutar.Controls.Add(this.lblResumenMontoProcesado);
			this.tabEjecutar.Controls.Add(this.lblResumenMontoDescuento);
			this.tabEjecutar.Controls.Add(this.tabResultado);
			this.tabEjecutar.Location = new System.Drawing.Point(4, 22);
			this.tabEjecutar.Name = "tabEjecutar";
			this.tabEjecutar.Size = new System.Drawing.Size(862, 494);
			this.tabEjecutar.TabIndex = 2;
			this.tabEjecutar.Text = "3. Ejecutar";
			// 
			// btnEjecutar
			// 
			this.btnEjecutar.BackColor = System.Drawing.Color.DarkRed;
			this.btnEjecutar.Enabled = false;
			this.btnEjecutar.ForeColor = System.Drawing.Color.White;
			this.btnEjecutar.Location = new System.Drawing.Point(8, 8);
			this.btnEjecutar.Name = "btnEjecutar";
			this.btnEjecutar.Size = new System.Drawing.Size(140, 28);
			this.btnEjecutar.TabIndex = 0;
			this.btnEjecutar.Text = "Ejecutar Cierre";
			this.btnEjecutar.Click += new System.EventHandler(this.btnEjecutar_Click);
			// 
			// lblResumenTotal
			// 
			this.lblResumenTotal.Location = new System.Drawing.Point(162, 10);
			this.lblResumenTotal.Name = "lblResumenTotal";
			this.lblResumenTotal.Size = new System.Drawing.Size(150, 18);
			this.lblResumenTotal.TabIndex = 1;
			this.lblResumenTotal.Text = "Total: 0";
			// 
			// lblResumenExitosos
			// 
			this.lblResumenExitosos.Location = new System.Drawing.Point(320, 10);
			this.lblResumenExitosos.Name = "lblResumenExitosos";
			this.lblResumenExitosos.Size = new System.Drawing.Size(120, 18);
			this.lblResumenExitosos.TabIndex = 2;
			this.lblResumenExitosos.Text = "Exitosas: 0";
			// 
			// lblResumenFallidos
			// 
			this.lblResumenFallidos.Location = new System.Drawing.Point(448, 10);
			this.lblResumenFallidos.Name = "lblResumenFallidos";
			this.lblResumenFallidos.Size = new System.Drawing.Size(120, 18);
			this.lblResumenFallidos.TabIndex = 3;
			this.lblResumenFallidos.Text = "Fallidas: 0";
			// 
			// lblResumenMontoProcesado
			// 
			this.lblResumenMontoProcesado.Location = new System.Drawing.Point(578, 10);
			this.lblResumenMontoProcesado.Name = "lblResumenMontoProcesado";
			this.lblResumenMontoProcesado.Size = new System.Drawing.Size(180, 18);
			this.lblResumenMontoProcesado.TabIndex = 4;
			this.lblResumenMontoProcesado.Text = "Monto recaudado: $0.00";
			// 
			// lblResumenMontoDescuento
			// 
			this.lblResumenMontoDescuento.Location = new System.Drawing.Point(162, 30);
			this.lblResumenMontoDescuento.Name = "lblResumenMontoDescuento";
			this.lblResumenMontoDescuento.Size = new System.Drawing.Size(250, 18);
			this.lblResumenMontoDescuento.TabIndex = 5;
			this.lblResumenMontoDescuento.Text = "Descuento total: $0.00";
			// 
			// tabResultado
			// 
			this.tabResultado.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabResultado.Controls.Add(this.tabExitosos);
			this.tabResultado.Controls.Add(this.tabFallidos);
			this.tabResultado.Location = new System.Drawing.Point(8, 52);
			this.tabResultado.Name = "tabResultado";
			this.tabResultado.SelectedIndex = 0;
			this.tabResultado.Size = new System.Drawing.Size(846, 434);
			this.tabResultado.TabIndex = 6;
			// 
			// tabExitosos
			// 
			this.tabExitosos.Controls.Add(this.lblFiltroExitosos);
			this.tabExitosos.Controls.Add(this.txtFiltroExitosos);
			this.tabExitosos.Controls.Add(this.gridExitosos);
			this.tabExitosos.Location = new System.Drawing.Point(0, 0);
			this.tabExitosos.Name = "tabExitosos";
			this.tabExitosos.Size = new System.Drawing.Size(838, 412);
			this.tabExitosos.TabIndex = 0;
			this.tabExitosos.Text = "Exitosas";
			// 
			// lblFiltroExitosos
			// 
			this.lblFiltroExitosos.Location = new System.Drawing.Point(2, 4);
			this.lblFiltroExitosos.Name = "lblFiltroExitosos";
			this.lblFiltroExitosos.Size = new System.Drawing.Size(45, 18);
			this.lblFiltroExitosos.TabIndex = 0;
			this.lblFiltroExitosos.Text = "Buscar:";
			// 
			// txtFiltroExitosos
			// 
			this.txtFiltroExitosos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFiltroExitosos.Location = new System.Drawing.Point(50, 2);
			this.txtFiltroExitosos.Name = "txtFiltroExitosos";
			this.txtFiltroExitosos.Size = new System.Drawing.Size(760, 20);
			this.txtFiltroExitosos.TabIndex = 1;
			this.txtFiltroExitosos.Text = "";
			this.txtFiltroExitosos.TextChanged += new System.EventHandler(this.txtFiltroExitosos_TextChanged);
			// 
			// gridExitosos
			// 
			this.gridExitosos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridExitosos.Location = new System.Drawing.Point(0, 28);
			this.gridExitosos.Name = "gridExitosos";
			this.gridExitosos.Size = new System.Drawing.Size(838, 384);
			this.gridExitosos.TabIndex = 2;
			// 
			// tabFallidos
			// 
			this.tabFallidos.Controls.Add(this.gridFallidos);
			this.tabFallidos.Location = new System.Drawing.Point(0, 0);
			this.tabFallidos.Name = "tabFallidos";
			this.tabFallidos.TabIndex = 1;
			this.tabFallidos.Text = "Fallidas";
			// 
			// gridFallidos
			// 
			this.gridFallidos.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridFallidos.Location = new System.Drawing.Point(0, 0);
			this.gridFallidos.Name = "gridFallidos";
			this.gridFallidos.Size = new System.Drawing.Size(200, 100);
			this.gridFallidos.TabIndex = 0;
			// 
			// statusBar
			// 
			this.statusBar.Location = new System.Drawing.Point(0, 0);
			this.statusBar.Name = "statusBar";
			this.statusBar.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
																																								 this.sbpMensaje});
			this.statusBar.ShowPanels = true;
			this.statusBar.TabIndex = 3;
			// 
			// sbpMensaje
			// 
			this.sbpMensaje.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring;
			this.sbpMensaje.Text = "Listo";
			this.sbpMensaje.Width = 84;
			// 
			// progressBar
			// 
			this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.progressBar.Location = new System.Drawing.Point(8, 622);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(870, 14);
			this.progressBar.TabIndex = 2;
			this.progressBar.Visible = false;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmCierreOperacionesDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(892, 646);
			this.Controls.Add(this.grpAutorizacion);
			this.Controls.Add(this.tabControl);
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.statusBar);
			this.Name = "frmCierreOperacionesDescuento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cierre de Operaciones con Descuento Especial";
			this.Load += new System.EventHandler(this.frmCierreOperacionesDescuento_Load);
			this.grpAutorizacion.ResumeLayout(false);
			this.tabControl.ResumeLayout(false);
			this.tabCargarExcel.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridExcel)).EndInit();
			this.tabValidar.ResumeLayout(false);
			this.tabValidacion.ResumeLayout(false);
			this.tabEncontrados.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridEncontrados)).EndInit();
			this.tabErrores.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridErrores)).EndInit();
			this.tabEjecutar.ResumeLayout(false);
			this.tabResultado.ResumeLayout(false);
			this.tabExitosos.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridExitosos)).EndInit();
			this.tabFallidos.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridFallidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
	}
}
