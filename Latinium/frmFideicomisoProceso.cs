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

namespace Latinium
{
	public class frmFideicomisoProceso : System.Windows.Forms.Form
	{
		// --- Panel superior fijo ---
		private System.Windows.Forms.GroupBox grpProceso;
		private System.Windows.Forms.Label lblEntidadDestino;
		private System.Windows.Forms.ComboBox cmbEntidadDestino;
		private System.Windows.Forms.Label lblFechaCorte;
		private System.Windows.Forms.DateTimePicker dtpFechaCorte;
		private System.Windows.Forms.Button btnDescargarPlantilla;
		private System.Windows.Forms.Button btnManual;

		// --- TabControl principal ---
		private System.Windows.Forms.TabControl tabControl;
		private System.Windows.Forms.TabPage tabCargarExcel;
		private System.Windows.Forms.TabPage tabPreparar;
		private System.Windows.Forms.TabPage tabVentaCartera;

		// --- Tab 1: Cargar Excel ---
		private System.Windows.Forms.Label lblArchivoExcel;
		private System.Windows.Forms.TextBox txtArchivoExcel;
		private System.Windows.Forms.Button btnBuscarExcel;
		private System.Windows.Forms.Button btnCargarExcel;
		private System.Windows.Forms.DataGrid gridExcel;
		private System.Windows.Forms.Label lblTotalExcel;

		// --- Tab 2: Preparar y Validar ---
		private System.Windows.Forms.Button btnPreparar;
		private System.Windows.Forms.Label lblContTotal;
		private System.Windows.Forms.Label lblContEncontrados;
		private System.Windows.Forms.Label lblContSinCoincidencia;
		private System.Windows.Forms.Label lblContYaProcesados;
		private System.Windows.Forms.TabControl tabValidacion;
		private System.Windows.Forms.TabPage tabEncontrados;
		private System.Windows.Forms.TabPage tabSinCoincidencia;
		private System.Windows.Forms.DataGrid gridEncontrados;
		private System.Windows.Forms.DataGrid gridSinCoincidencia;

		// --- Tab 3: Ejecutar Venta ---
		private int numeroLoteCalculado = 0;
		private System.Windows.Forms.Button btnEjecutarVenta;
		private System.Windows.Forms.Label lblResumenTotal;
		private System.Windows.Forms.Label lblResumenActualizados;
		private System.Windows.Forms.Label lblResumenFallidos;
		private System.Windows.Forms.TabControl tabResultadoVenta;
		private System.Windows.Forms.TabPage tabActualizados;
		private System.Windows.Forms.TabPage tabFallidos;
		private System.Windows.Forms.DataGrid gridActualizados;
		private System.Windows.Forms.DataGrid gridFallidos;

		// --- Barra inferior ---
		private System.Windows.Forms.StatusBar statusBar;
		private System.Windows.Forms.StatusBarPanel sbpMensaje;
		private System.Windows.Forms.ProgressBar progressBar;

		// --- Seguridad ---
		private Acceso miAcceso;
		private C1.Data.C1DataSet cdsSeteoF;

		// --- Constantes de negocio ---
		private const int TipoCarteFideicomiso = 1;

		// --- Estado del flujo ---
		private DataTable dtExcel;
		private bool excelCargado        = false;
		private bool preparacionHecha    = false;
		private int  encontradosCount    = 0;
		private int  sinCoincidenciaCount = 0;
		private int  yaProcesadosCount    = 0;

		// --- Búsqueda en grids ---
		private System.Windows.Forms.Label   lblBuscarExcel;
		private System.Windows.Forms.TextBox txtBuscarExcel;
		private System.Windows.Forms.Label   lblBuscarEncontrados;
		private System.Windows.Forms.TextBox txtBuscarEncontrados;
		private System.Windows.Forms.Label   lblBuscarSinCoinc;
		private System.Windows.Forms.TextBox txtBuscarSinCoinc;
		private System.Windows.Forms.Label   lblBuscarActualizados;
		private System.Windows.Forms.TextBox txtBuscarActualizados;
		private System.Windows.Forms.Label   lblBuscarFallidos;
		private System.Windows.Forms.TextBox txtBuscarFallidos;

		// --- DataViews para filtrado ---
		private DataView dvExcel;
		private DataView dvEncontrados;
		private DataView dvSinCoincidencia;
		private DataView dvActualizados;
		private DataView dvFallidos;

		private System.ComponentModel.Container components = null;

		public frmFideicomisoProceso()
		{
			InitializeComponent();
			CargarEntidades();
			ActualizarBotones();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null) components.Dispose();
			}
			base.Dispose(disposing);
		}

		private SqlConnection AbrirConexion()
		{
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			SqlConnection cn = new SqlConnection(strConn);
			cn.Open();
			return cn;
		}

		// ================================================================
		//  CARGA INICIAL DE ENTIDADES
		// ================================================================

		private void CargarEntidades()
		{
			try
			{
				Funcion miFuncion = new Funcion();
				string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
				SqlDataAdapter da = new SqlDataAdapter(
					"SELECT idEntidadFinanciera, Nombre FROM Cre_EntidadFinanciera " +
					"WHERE Activo = 1 AND idCre_TipoCarte = " + TipoCarteFideicomiso +
					" ORDER BY Nombre", strConn);
				DataTable dt = new DataTable();
				da.Fill(dt);

				cmbEntidadDestino.DataSource    = dt;
				cmbEntidadDestino.DisplayMember = "Nombre";
				cmbEntidadDestino.ValueMember   = "idEntidadFinanciera";
				cmbEntidadDestino.SelectedIndex = -1;
			}
			catch (Exception ex)
			{
				SetStatus("Advertencia al cargar entidades: " + ex.Message);
			}
		}

		private void cmbEntidadDestino_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (cmbEntidadDestino.SelectedItem == null) return;
			string nombre = ((DataRowView)cmbEntidadDestino.SelectedItem)["Nombre"].ToString();
			this.Text = "Venta de Cartera Fideicomiso  ->  " + nombre;
			CargarProximoLote();
		}

		private void CargarProximoLote()
		{
			if (cmbEntidadDestino.SelectedValue == null) return;
			try
			{
				int idEntidad = Convert.ToInt32(cmbEntidadDestino.SelectedValue);
				Funcion miFuncion = new Funcion();
				string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
				SqlConnection cn = new SqlConnection(strConn);
				cn.Open();
				SqlCommand cmd = new SqlCommand(
					"SELECT ISNULL(MAX(NumeroLote), 0) + 1 " +
					"FROM Compra " +
					"WHERE idTipoFactura = 1 AND idEntidadFinanciera = @idEntidad", cn);
				cmd.Parameters.Add("@idEntidad", SqlDbType.Int).Value = idEntidad;
				object result = cmd.ExecuteScalar();
				cn.Close();
				numeroLoteCalculado = (result != null) ? Convert.ToInt32(result) : 1;
			}
			catch
			{
				numeroLoteCalculado = 0;
			}
		}

		// ================================================================
		//  PLANTILLA EXCEL
		// ================================================================

		private void btnManual_Click(object sender, System.EventArgs e)
		{
			try
			{
				Funcion miFuncion = new Funcion();
				string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
				SqlConnection cn = new SqlConnection(strConn);
				cn.Open();
				SqlCommand cmd = new SqlCommand(
					"SELECT TOP 1 Url FROM DocumentacionConfiguracion " +
					"WHERE Clave = 'ManualFideicomiso' AND Activo = 1", cn);
				object result = cmd.ExecuteScalar();
				cn.Close();

				if (result != null && result != DBNull.Value)
				{
					System.Diagnostics.Process.Start(result.ToString());
				}
				else
				{
					MessageBox.Show(
						"No se encontro la URL del manual en DocumentacionConfiguracion.\n" +
						"Ejecute el script CrearDocumentacionConfiguracion.sql en la base de datos.",
						"Manual no configurado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al obtener URL del manual:\n" + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnDescargarPlantilla_Click(object sender, System.EventArgs e)
		{
			SaveFileDialog dlg = new SaveFileDialog();
			dlg.Title    = "Guardar plantilla";
			dlg.Filter   = "Excel 97-2003 (*.xls)|*.xls";
			dlg.FileName = "Plantilla_Fideicomiso.xls";
			if (dlg.ShowDialog() != DialogResult.OK) return;
			try
			{
				GenerarPlantilla(dlg.FileName);
				SetStatus("Plantilla guardada: " + dlg.FileName);
				MessageBox.Show("Plantilla guardada correctamente en:\n" + dlg.FileName,
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
			sw.WriteLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
			sw.WriteLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
			sw.WriteLine(" <Styles>");
			sw.WriteLine("  <Style ss:ID=\"sTitulo\">");
			sw.WriteLine("   <Font ss:Bold=\"1\" ss:Size=\"12\" ss:Color=\"#FFFFFF\"/>");
			sw.WriteLine("   <Interior ss:Color=\"#003366\" ss:Pattern=\"Solid\"/>");
			sw.WriteLine("  </Style>");
			sw.WriteLine("  <Style ss:ID=\"sEncabezado\">");
			sw.WriteLine("   <Font ss:Bold=\"1\"/>");
			sw.WriteLine("   <Interior ss:Color=\"#C0C0C0\" ss:Pattern=\"Solid\"/>");
			sw.WriteLine("  </Style>");
			sw.WriteLine("  <Style ss:ID=\"sNota\">");
			sw.WriteLine("   <Font ss:Italic=\"1\" ss:Color=\"#CC0000\"/>");
			sw.WriteLine("  </Style>");
			sw.WriteLine(" </Styles>");
			sw.WriteLine(" <Worksheet ss:Name=\"Fideicomiso\">");
			sw.WriteLine("  <Table ss:DefaultColumnWidth=\"80\">");
			sw.WriteLine("   <Column ss:Index=\"1\" ss:Width=\"35\"/>");
			sw.WriteLine("   <Column ss:Index=\"2\" ss:Width=\"110\"/>");
			sw.WriteLine("   <Column ss:Index=\"3\" ss:Width=\"220\"/>");
			sw.WriteLine("   <Column ss:Index=\"4\" ss:Width=\"130\"/>");
			sw.WriteLine("   <Column ss:Index=\"5\" ss:Width=\"110\"/>");
			sw.WriteLine("   <Row ss:Height=\"22\">");
			sw.WriteLine("    <Cell ss:Index=\"2\" ss:StyleID=\"sTitulo\"><Data ss:Type=\"String\">PLANTILLA  -  VENTA CARTERA FIDEICOMISO</Data></Cell>");
			sw.WriteLine("   </Row>");
			sw.WriteLine("   <Row>");
			sw.WriteLine("    <Cell ss:Index=\"2\" ss:StyleID=\"sNota\"><Data ss:Type=\"String\">INSTRUCCIONES: No eliminar ni modificar las filas 1 al 7. Los datos deben comenzar EXACTAMENTE desde la fila 8.</Data></Cell>");
			sw.WriteLine("   </Row>");
			sw.WriteLine("   <Row>");
			sw.WriteLine("    <Cell ss:Index=\"2\" ss:StyleID=\"sNota\"><Data ss:Type=\"String\">Columna B: Cedula de identidad del cliente (maximo 10 caracteres).</Data></Cell>");
			sw.WriteLine("   </Row>");
			sw.WriteLine("   <Row>");
			sw.WriteLine("    <Cell ss:Index=\"2\" ss:StyleID=\"sNota\"><Data ss:Type=\"String\">Columna D: Numero de operacion del credito (maximo 9 caracteres).</Data></Cell>");
			sw.WriteLine("   </Row>");
			sw.WriteLine("   <Row>");
			sw.WriteLine("    <Cell ss:Index=\"2\" ss:StyleID=\"sNota\"><Data ss:Type=\"String\">Columna E: Fecha de la operacion (formato DD/MM/AAAA). El sistema tomara la fecha mas reciente como Fecha de Corte.</Data></Cell>");
			sw.WriteLine("   </Row>");
			sw.WriteLine("   <Row/>");
			sw.WriteLine("   <Row ss:Height=\"18\">");
			sw.WriteLine("    <Cell ss:StyleID=\"sEncabezado\"><Data ss:Type=\"String\">N</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEncabezado\"><Data ss:Type=\"String\">CEDULA</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEncabezado\"><Data ss:Type=\"String\">NOMBRES</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEncabezado\"><Data ss:Type=\"String\">NUMERO_OPERACION</Data></Cell>");
			sw.WriteLine("    <Cell ss:StyleID=\"sEncabezado\"><Data ss:Type=\"String\">FECHA_OPERACION</Data></Cell>");
			sw.WriteLine("   </Row>");
			for (int i = 1; i <= 10; i++)
			{
				sw.WriteLine("   <Row>");
				sw.WriteLine("    <Cell><Data ss:Type=\"Number\">" + i + "</Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
				sw.WriteLine("   </Row>");
			}
			sw.WriteLine("  </Table>");
			sw.WriteLine(" </Worksheet>");
			sw.WriteLine("</Workbook>");
			sw.Close();
		}

		// ================================================================
		//  TAB 1 - CARGAR EXCEL
		// ================================================================

		private void btnBuscarExcel_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog dlg = new OpenFileDialog();
			dlg.Title  = "Seleccionar archivo CSV";
			dlg.Filter = "CSV (*.csv)|*.csv|Todos|*.*";
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

				txtBuscarExcel.Text  = "";
				dvExcel              = new DataView(dtExcel);
				gridExcel.DataSource = dvExcel;
				lblTotalExcel.Text   = "Total leidos: " + dtExcel.Rows.Count;
				excelCargado         = dtExcel.Rows.Count > 0;
				preparacionHecha     = false;
				yaProcesadosCount    = 0;

				SetStatus("Excel cargado: " + dtExcel.Rows.Count + " registros.");

				ActualizarBotones();
			}
			catch (Exception ex)
			{
				SetStatus("Error: " + ex.Message);
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

				OleDbDataAdapter da = new OleDbDataAdapter(
					"SELECT F2, F4, F5 FROM [" + hoja + "]", oleConn);
				DataTable dtTemp = new DataTable();
				da.Fill(dtTemp);

				dtExcel = new DataTable();
				dtExcel.Columns.Add("Cedula",          typeof(string));
				dtExcel.Columns.Add("NumeroOperacion", typeof(string));
				dtExcel.Columns.Add("FechaOperacion",  typeof(DateTime));

				// Primeras 7 filas = encabezados; datos desde fila 8 (indice 7)
				for (int i = 7; i < dtTemp.Rows.Count; i++)
				{
					string cedula = (dtTemp.Rows[i][0] == DBNull.Value)
						? "" : dtTemp.Rows[i][0].ToString().Trim();
					string numero = (dtTemp.Rows[i][1] == DBNull.Value)
						? "" : dtTemp.Rows[i][1].ToString().Trim();
					if (cedula.Length > 10) cedula = cedula.Substring(0, 10);
					if (cedula != "") cedula = cedula.PadLeft(10, '0');
					if (numero.Length > 9) numero = numero.Substring(0, 9);
					if (numero != "") numero = numero.PadLeft(9, '0');
					if (cedula != "" && numero != "")
					{
						DateTime fechaOp = DateTime.MinValue;
						if (dtTemp.Rows[i][2] != DBNull.Value)
						{
							try { fechaOp = Convert.ToDateTime(dtTemp.Rows[i][2]); }
							catch { }
						}
						DataRow nr = dtExcel.NewRow();
						nr["Cedula"]          = cedula;
						nr["NumeroOperacion"] = numero;
						nr["FechaOperacion"]  = (fechaOp == DateTime.MinValue)
							? (object)DBNull.Value : fechaOp;
						dtExcel.Rows.Add(nr);
					}
				}
			}
			finally
			{
				oleConn.Close();
			}
		}

		private void LeerCSV(string archivo)
		{
			dtExcel = new DataTable();
			dtExcel.Columns.Add("Cedula",          typeof(string));
			dtExcel.Columns.Add("NumeroOperacion", typeof(string));
			dtExcel.Columns.Add("FechaOperacion",  typeof(DateTime));

			StreamReader sr = new StreamReader(archivo, Encoding.Default);
			int lineaNum = 0;
			string linea;
			while ((linea = sr.ReadLine()) != null)
			{
				lineaNum++;
				if (lineaNum < 8) continue;
				string[] campos = linea.Split(',');
				if (campos.Length >= 4)
				{
					string cedula = campos[1].Trim().Replace("\"", "");
					string numero = campos[3].Trim().Replace("\"", "");
					if (cedula.Length > 10) cedula = cedula.Substring(0, 10);
					if (cedula != "") cedula = cedula.PadLeft(10, '0');
					if (numero.Length > 9) numero = numero.Substring(0, 9);
					if (numero != "") numero = numero.PadLeft(9, '0');
					if (cedula != "" && numero != "")
					{
						DateTime fechaOp = DateTime.MinValue;
						if (campos.Length >= 5)
						{
							string fechaStr = campos[4].Trim().Replace("\"", "");
							try { fechaOp = DateTime.Parse(fechaStr); }
							catch { }
						}
						DataRow nr = dtExcel.NewRow();
						nr["Cedula"]          = cedula;
						nr["NumeroOperacion"] = numero;
						nr["FechaOperacion"]  = (fechaOp == DateTime.MinValue)
							? (object)DBNull.Value : fechaOp;
						dtExcel.Rows.Add(nr);
					}
				}
			}
			sr.Close();
		}

		// ================================================================
		//  TAB 2 - PREPARAR Y VALIDAR
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
				SetStatus("Limpiando temporalfideicomiso...");
				MostrarProgress(true);
				int total = dtExcel.Rows.Count;
				progressBar.Maximum = total + 20;
				progressBar.Value   = 0;

				cn = AbrirConexion();

				// 1. Limpiar tabla
				new SqlCommand("DELETE FROM temporalfideicomiso", cn).ExecuteNonQuery();
				progressBar.Value = 5;
				Application.DoEvents();

				// 2. Insertar registros del Excel en lotes de 100
				SetStatus("Cargando registros en temporalfideicomiso...");
				int desde    = 0;
				int loteSize = 100;
				while (desde < total)
				{
					int hasta = desde + loteSize;
					if (hasta > total) hasta = total;

					StringBuilder sb = new StringBuilder();
					sb.Append("INSERT INTO temporalfideicomiso (cedula, numero, idcompra) VALUES ");
					for (int i = desde; i < hasta; i++)
					{
						string ced = dtExcel.Rows[i]["Cedula"].ToString().Replace("'", "''");
						string num = dtExcel.Rows[i]["NumeroOperacion"].ToString().Replace("'", "''");
						if (ced.Length > 10) ced = ced.Substring(0, 10);
						if (num.Length > 9)  num = num.Substring(0, 9);
						ced = ced.PadLeft(10, '0');
						num = num.PadLeft(9,  '0');
						if (i > desde) sb.Append(",");
						sb.Append("('" + ced + "','" + num + "', NULL)");
					}
					new SqlCommand(sb.ToString(), cn).ExecuteNonQuery();
					desde = hasta;
					progressBar.Value = 5 + desde;
					Application.DoEvents();
				}

				// 3. Cruzar con Compra para obtener idcompra
				SetStatus("Cruzando con tabla Compra...");
				SqlCommand cmdUpdate = new SqlCommand(
					"UPDATE temporalfideicomiso " +
					"SET idcompra = c.idcompra " +
					"FROM temporalfideicomiso t " +
					"LEFT JOIN Compra c ON t.cedula = c.Ruc AND t.numero = c.Numero " +
					"    AND c.idTipoFactura = 1 AND c.Borrar = 0 AND c.Estado = 10 " +
					"    AND c.idFormaPago = 9", cn);
				cmdUpdate.CommandTimeout = 120;
				cmdUpdate.ExecuteNonQuery();
				progressBar.Value = total + 15;
				Application.DoEvents();

				// Fecha de corte = MIN(Fecha) de los registros encontrados en Compra
				SqlCommand cmdFecha = new SqlCommand(
					"SELECT MIN(c.Fecha) " +
					"FROM temporalfideicomiso t " +
					"INNER JOIN Compra c ON t.idcompra = c.idCompra", cn);
				object fechaMin = cmdFecha.ExecuteScalar();
				if (fechaMin != null && fechaMin != DBNull.Value)
					dtpFechaCorte.Value = Convert.ToDateTime(fechaMin).Date;

				// 4. Mostrar ENCONTRADOS
				SqlDataAdapter daEnc = new SqlDataAdapter(
					"SELECT t.cedula AS Cedula, t.numero AS NumeroOperacion, " +
					"       c.Numero AS NumeroCompra, " +
					"       ef.Nombre AS Entidad, c.EstadoSolNC, c.SaldoAnticipo, " +
					"       CASE WHEN c.idEntidadFinanciera <> 1 " +
					"            THEN 'SI - ' + ISNULL(ef.Nombre, CAST(c.idEntidadFinanciera AS varchar(10))) " +
					"            ELSE 'No' END AS YaProcesado " +
					"FROM temporalfideicomiso t " +
					"INNER JOIN Compra c ON t.idcompra = c.idCompra " +
					"LEFT JOIN Cre_EntidadFinanciera ef ON c.idEntidadFinanciera = ef.idEntidadFinanciera " +
					"ORDER BY t.cedula", cn);
				DataTable dtEnc = new DataTable();
				daEnc.Fill(dtEnc);

				// 5. Mostrar SIN COINCIDENCIA con motivo detallado
				SqlCommand cmdDiag = new SqlCommand(
					"SELECT t.cedula AS Cedula, t.numero AS NumeroOperacion, " +
					"  CASE " +
					"    WHEN c.idCompra IS NULL    THEN 'No existe en Compra (cedula/numero no encontrado)' " +
					"    WHEN c.Borrar  = 1         THEN 'Excluido: Borrar=1' " +
					"    WHEN c.idTipoFactura <> 1  THEN 'Excluido: TipoFactura=' + CAST(c.idTipoFactura AS varchar(5)) " +
					"    WHEN c.Estado  <> 10       THEN 'Excluido: Estado=' + CAST(c.Estado AS varchar(5)) " +
					"    WHEN c.idFormaPago <> 9    THEN 'Excluido: FormaPago=' + CAST(c.idFormaPago AS varchar(5)) " +
					"    ELSE 'Sin coincidencia' " +
					"  END AS Motivo " +
					"FROM temporalfideicomiso t " +
					"LEFT JOIN Compra c ON t.cedula = c.Ruc AND t.numero = c.Numero " +
					"WHERE t.idcompra IS NULL " +
					"ORDER BY t.cedula", cn);
				SqlDataAdapter daSin = new SqlDataAdapter(cmdDiag);
				DataTable dtSin = new DataTable();
				daSin.Fill(dtSin);

				txtBuscarEncontrados.Text       = "";
				dvEncontrados                   = new DataView(dtEnc);
				gridEncontrados.DataSource      = dvEncontrados;

				txtBuscarSinCoinc.Text          = "";
				dvSinCoincidencia               = new DataView(dtSin);
				gridSinCoincidencia.DataSource  = dvSinCoincidencia;
				encontradosCount     = dtEnc.Rows.Count;
				sinCoincidenciaCount = dtSin.Rows.Count;

				SqlCommand cmdYaProc = new SqlCommand(
					"SELECT COUNT(*) FROM temporalfideicomiso t " +
					"INNER JOIN Compra c ON t.idcompra = c.idCompra " +
					"WHERE c.idEntidadFinanciera <> 1", cn);
				yaProcesadosCount = Convert.ToInt32(cmdYaProc.ExecuteScalar());

				lblContTotal.Text           = "Total Excel: "      + total;
				lblContEncontrados.Text     = "Encontrados: "      + dtEnc.Rows.Count;
				lblContSinCoincidencia.Text = "Sin coincidencia: " + dtSin.Rows.Count;
				lblContYaProcesados.Text    = "Ya procesados: "    + yaProcesadosCount;
				lblContYaProcesados.ForeColor = (yaProcesadosCount > 0)
					? System.Drawing.Color.DarkRed
					: System.Drawing.Color.Green;

				preparacionHecha = true;
				ActualizarBotones();
				SetStatus("Preparacion completa.  Encontrados: " + dtEnc.Rows.Count
					+ "  |  Sin coincidencia: " + dtSin.Rows.Count);
				progressBar.Value = total + 20;
			}
			catch (Exception ex)
			{
				SetStatus("Error en preparacion: " + ex.Message);
				MessageBox.Show("Error al preparar datos:\n" + ex.Message,
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
		//  TAB 3 - EJECUTAR VENTA DE CARTERA
		// ================================================================

		private void btnEjecutarVenta_Click(object sender, System.EventArgs e)
		{
			if (!preparacionHecha || encontradosCount == 0)
			{
				MessageBox.Show("Debe completar el paso 'Preparar y Validar' y tener registros encontrados antes de ejecutar la venta.",
					"Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			if (sinCoincidenciaCount > 0)
			{
				MessageBox.Show(
					"No se puede ejecutar la venta.\n\n" +
					"Existen " + sinCoincidenciaCount + " registro(s) sin coincidencia en Compra.\n\n" +
					"Todos los registros del Excel deben existir en el sistema antes de proceder.\n" +
					"Revise la pestana '2. Preparar y Validar' > SIN COINCIDENCIA para ver el detalle.",
					"Registros sin coincidencia",
					MessageBoxButtons.OK, MessageBoxIcon.Stop);
				tabControl.SelectedTab = tabPreparar;
				return;
			}
			if (cmbEntidadDestino.SelectedValue == null)
			{
				MessageBox.Show("Seleccione la Entidad Destino en el panel superior.", "Aviso",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			if (numeroLoteCalculado <= 0)
				CargarProximoLote();

			if (numeroLoteCalculado <= 0)
			{
				MessageBox.Show("No se pudo calcular el numero de lote. Verifique la conexion.", "Aviso",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			int numeroLote = numeroLoteCalculado;

			int procesables = encontradosCount - yaProcesadosCount;
			string nombreDestino = ((DataRowView)cmbEntidadDestino.SelectedItem)["Nombre"].ToString();
			string msgOmitidos = (yaProcesadosCount > 0)
				? "  Se omitiran     : " + yaProcesadosCount + " registro(s) ya asignados a otra entidad\n"
				: "";
			if (MessageBox.Show(
				"Confirma la ejecucion de la Venta de Cartera?\n\n" +
				"  Entidad destino  : " + nombreDestino + "\n" +
				"  Numero de lote   : " + numeroLote + "\n\n" +
				"  Se actualizaran  : " + procesables + " registro(s) en la tabla Compra.\n" +
				msgOmitidos,
				"Confirmar Venta",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
				return;

			SqlConnection cn = null;
			try
			{
				Cursor = Cursors.WaitCursor;
				SetStatus("Ejecutando Venta de Cartera...");
				MostrarProgress(true);
				progressBar.Maximum = 100;
				progressBar.Value   = 20;

				int idDestino = Convert.ToInt32(cmbEntidadDestino.SelectedValue);

				cn = AbrirConexion();

				// UPDATE Compra para los registros encontrados
				SqlCommand cmdVenta = new SqlCommand(
					"UPDATE Compra " +
					"SET Bloquear = 1, CapturaLote = 1, NumeroLote = @NumeroLote, " +
					"    idEntidadFinanciera = @idEntidadDestino, " +
					"    FechaTransitoLote = GETDATE(), FechaVenta = GETDATE(), " +
					"    NumeroDeOperacion = Ruc + '-' + NUMERO, Confirmado = 1 " +
					"WHERE idEntidadFinanciera = 1 " +
					"  AND idCompra IN " +
					"  (SELECT idCompra FROM temporalfideicomiso WHERE idCompra IS NOT NULL)", cn);
				cmdVenta.Parameters.Add("@NumeroLote",       SqlDbType.Int).Value = numeroLote;
				cmdVenta.Parameters.Add("@idEntidadDestino", SqlDbType.Int).Value = idDestino;
				cmdVenta.CommandTimeout = 300;
				cmdVenta.ExecuteNonQuery();

				progressBar.Value = 60;
				Application.DoEvents();

				// Mostrar ACTUALIZADOS
				SqlDataAdapter daAct = new SqlDataAdapter(
					"SELECT t.cedula AS Cedula, t.numero AS NumeroOperacion, " +
					"       c.Numero AS NumeroCompra, c.NumeroLote, " +
					"       ef.Nombre AS Entidad, c.NumeroDeOperacion, c.FechaVenta " +
					"FROM temporalfideicomiso t " +
					"INNER JOIN Compra c ON t.idcompra = c.idCompra " +
					"LEFT JOIN Cre_EntidadFinanciera ef ON c.idEntidadFinanciera = ef.idEntidadFinanciera " +
					"ORDER BY t.cedula", cn);
				DataTable dtAct = new DataTable();
				daAct.Fill(dtAct);

				// Mostrar SIN COINCIDENCIA (fallidos)
				SqlDataAdapter daFall = new SqlDataAdapter(
					"SELECT cedula AS Cedula, numero AS NumeroOperacion, " +
					"       'Sin coincidencia en Compra' AS Motivo " +
					"FROM temporalfideicomiso " +
					"WHERE idcompra IS NULL " +
					"ORDER BY cedula", cn);
				DataTable dtFall = new DataTable();
				daFall.Fill(dtFall);

				progressBar.Value = 90;

				txtBuscarActualizados.Text  = "";
				dvActualizados              = new DataView(dtAct);
				gridActualizados.DataSource = dvActualizados;

				txtBuscarFallidos.Text      = "";
				dvFallidos                  = new DataView(dtFall);
				gridFallidos.DataSource     = dvFallidos;

				int total = dtAct.Rows.Count + dtFall.Rows.Count;
				lblResumenTotal.Text        = "Total: "           + total;
				lblResumenActualizados.Text = "Actualizados: "    + dtAct.Rows.Count;
				lblResumenFallidos.Text     = "Sin coincidencia: "+ dtFall.Rows.Count;

				string msg = "Venta completada para " + nombreDestino
					+ ".  Actualizados: " + dtAct.Rows.Count
					+ "  |  Sin coincidencia: " + dtFall.Rows.Count;
				SetStatus(msg);
				MessageBox.Show(msg, "Venta de Cartera",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				SetStatus("Error en venta: " + ex.Message);
				MessageBox.Show("Error al ejecutar venta:\n" + ex.Message,
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
		//  HELPERS
		// ================================================================

		private void SetStatus(string msg)
		{
			sbpMensaje.Text = msg;
			Application.DoEvents();
		}

		private void MostrarProgress(bool visible)
		{
			progressBar.Visible = visible;
			if (!visible) progressBar.Value = 0;
		}

		private void ActualizarBotones()
		{
			btnPreparar.Enabled      = excelCargado;
			btnEjecutarVenta.Enabled = preparacionHecha && (encontradosCount - yaProcesadosCount) > 0;
		}

		private void AplicarFiltro(DataView dv, string texto, string[] cols)
		{
			if (dv == null) return;
			try
			{
				if (texto.Trim() == "")
				{
					dv.RowFilter = "";
				}
				else
				{
					string t = texto.Trim().Replace("'", "''");
					StringBuilder sb = new StringBuilder();
					foreach (string col in cols)
					{
						if (sb.Length > 0) sb.Append(" OR ");
						sb.Append(col + " LIKE '%" + t + "%'");
					}
					dv.RowFilter = sb.ToString();
				}
			}
			catch { }
		}

		private void txtBuscarExcel_TextChanged(object sender, System.EventArgs e)
		{
			AplicarFiltro(dvExcel, txtBuscarExcel.Text,
				new string[] { "Cedula", "NumeroOperacion" });
		}

		private void txtBuscarEncontrados_TextChanged(object sender, System.EventArgs e)
		{
			AplicarFiltro(dvEncontrados, txtBuscarEncontrados.Text,
				new string[] { "Cedula", "NumeroOperacion", "Entidad" });
		}

		private void txtBuscarSinCoinc_TextChanged(object sender, System.EventArgs e)
		{
			AplicarFiltro(dvSinCoincidencia, txtBuscarSinCoinc.Text,
				new string[] { "Cedula", "NumeroOperacion", "Motivo" });
		}

		private void txtBuscarActualizados_TextChanged(object sender, System.EventArgs e)
		{
			AplicarFiltro(dvActualizados, txtBuscarActualizados.Text,
				new string[] { "Cedula", "NumeroOperacion", "Entidad", "NumeroDeOperacion" });
		}

		private void txtBuscarFallidos_TextChanged(object sender, System.EventArgs e)
		{
			AplicarFiltro(dvFallidos, txtBuscarFallidos.Text,
				new string[] { "Cedula", "NumeroOperacion", "Motivo" });
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmFideicomisoProceso_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320702");
			if (!Funcion.Permiso("2230", cdsSeteoF))
			{
				MessageBox.Show("No tiene acceso a Venta de Cartera Fideicomiso", "Point Technology",
					MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		}

		// ================================================================
		//  InitializeComponent
		// ================================================================

		#region Windows Form Designer generated code

		private void InitializeComponent()
		{
			this.grpProceso = new System.Windows.Forms.GroupBox();
			this.lblEntidadDestino = new System.Windows.Forms.Label();
			this.cmbEntidadDestino = new System.Windows.Forms.ComboBox();
			this.lblFechaCorte = new System.Windows.Forms.Label();
			this.dtpFechaCorte = new System.Windows.Forms.DateTimePicker();
			this.btnDescargarPlantilla = new System.Windows.Forms.Button();
			this.btnManual = new System.Windows.Forms.Button();
			this.tabControl = new System.Windows.Forms.TabControl();
			this.tabCargarExcel = new System.Windows.Forms.TabPage();
			this.lblArchivoExcel = new System.Windows.Forms.Label();
			this.txtArchivoExcel = new System.Windows.Forms.TextBox();
			this.btnBuscarExcel = new System.Windows.Forms.Button();
			this.btnCargarExcel = new System.Windows.Forms.Button();
			this.lblTotalExcel = new System.Windows.Forms.Label();
			this.gridExcel = new System.Windows.Forms.DataGrid();
			this.tabPreparar = new System.Windows.Forms.TabPage();
			this.btnPreparar = new System.Windows.Forms.Button();
			this.lblContTotal = new System.Windows.Forms.Label();
			this.lblContEncontrados = new System.Windows.Forms.Label();
			this.lblContSinCoincidencia = new System.Windows.Forms.Label();
			this.lblContYaProcesados = new System.Windows.Forms.Label();
			this.tabValidacion = new System.Windows.Forms.TabControl();
			this.tabEncontrados = new System.Windows.Forms.TabPage();
			this.gridEncontrados = new System.Windows.Forms.DataGrid();
			this.tabSinCoincidencia = new System.Windows.Forms.TabPage();
			this.gridSinCoincidencia = new System.Windows.Forms.DataGrid();
			this.tabVentaCartera = new System.Windows.Forms.TabPage();
			this.btnEjecutarVenta = new System.Windows.Forms.Button();
			this.lblResumenTotal = new System.Windows.Forms.Label();
			this.lblResumenActualizados = new System.Windows.Forms.Label();
			this.lblResumenFallidos = new System.Windows.Forms.Label();
			this.tabResultadoVenta = new System.Windows.Forms.TabControl();
			this.tabActualizados = new System.Windows.Forms.TabPage();
			this.gridActualizados = new System.Windows.Forms.DataGrid();
			this.tabFallidos = new System.Windows.Forms.TabPage();
			this.gridFallidos = new System.Windows.Forms.DataGrid();
			this.lblBuscarExcel = new System.Windows.Forms.Label();
			this.txtBuscarExcel = new System.Windows.Forms.TextBox();
			this.lblBuscarEncontrados = new System.Windows.Forms.Label();
			this.txtBuscarEncontrados = new System.Windows.Forms.TextBox();
			this.lblBuscarSinCoinc = new System.Windows.Forms.Label();
			this.txtBuscarSinCoinc = new System.Windows.Forms.TextBox();
			this.lblBuscarActualizados = new System.Windows.Forms.Label();
			this.txtBuscarActualizados = new System.Windows.Forms.TextBox();
			this.lblBuscarFallidos = new System.Windows.Forms.Label();
			this.txtBuscarFallidos = new System.Windows.Forms.TextBox();
			this.statusBar = new System.Windows.Forms.StatusBar();
			this.sbpMensaje = new System.Windows.Forms.StatusBarPanel();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.grpProceso.SuspendLayout();
			this.tabControl.SuspendLayout();
			this.tabCargarExcel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridExcel)).BeginInit();
			this.tabPreparar.SuspendLayout();
			this.tabValidacion.SuspendLayout();
			this.tabEncontrados.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridEncontrados)).BeginInit();
			this.tabSinCoincidencia.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridSinCoincidencia)).BeginInit();
			this.tabVentaCartera.SuspendLayout();
			this.tabResultadoVenta.SuspendLayout();
			this.tabActualizados.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridActualizados)).BeginInit();
			this.tabFallidos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridFallidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grpProceso
			// 
			this.grpProceso.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grpProceso.Controls.Add(this.lblEntidadDestino);
			this.grpProceso.Controls.Add(this.cmbEntidadDestino);
			this.grpProceso.Controls.Add(this.lblFechaCorte);
			this.grpProceso.Controls.Add(this.dtpFechaCorte);
			this.grpProceso.Controls.Add(this.btnDescargarPlantilla);
			this.grpProceso.Controls.Add(this.btnManual);
			this.grpProceso.Location = new System.Drawing.Point(4, 4);
			this.grpProceso.Name = "grpProceso";
			this.grpProceso.Size = new System.Drawing.Size(876, 50);
			this.grpProceso.TabIndex = 0;
			this.grpProceso.TabStop = false;
			this.grpProceso.Text = "Parametros del Proceso";
			// 
			// lblEntidadDestino
			// 
			this.lblEntidadDestino.Location = new System.Drawing.Point(8, 20);
			this.lblEntidadDestino.Name = "lblEntidadDestino";
			this.lblEntidadDestino.Size = new System.Drawing.Size(48, 16);
			this.lblEntidadDestino.TabIndex = 0;
			this.lblEntidadDestino.Text = "Entidad:";
			// 
			// cmbEntidadDestino
			// 
			this.cmbEntidadDestino.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbEntidadDestino.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
			this.cmbEntidadDestino.Location = new System.Drawing.Point(60, 16);
			this.cmbEntidadDestino.Name = "cmbEntidadDestino";
			this.cmbEntidadDestino.Size = new System.Drawing.Size(250, 23);
			this.cmbEntidadDestino.TabIndex = 1;
			this.cmbEntidadDestino.SelectedIndexChanged += new System.EventHandler(this.cmbEntidadDestino_SelectedIndexChanged);
			// 
			// lblFechaCorte
			// 
			this.lblFechaCorte.Location = new System.Drawing.Point(322, 20);
			this.lblFechaCorte.Name = "lblFechaCorte";
			this.lblFechaCorte.Size = new System.Drawing.Size(76, 16);
			this.lblFechaCorte.TabIndex = 2;
			this.lblFechaCorte.Text = "Fecha Corte:";
			// 
			// dtpFechaCorte
			// 
			this.dtpFechaCorte.Enabled = false;
			this.dtpFechaCorte.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpFechaCorte.Location = new System.Drawing.Point(402, 16);
			this.dtpFechaCorte.Name = "dtpFechaCorte";
			this.dtpFechaCorte.Size = new System.Drawing.Size(120, 20);
			this.dtpFechaCorte.TabIndex = 3;
			// 
			// btnDescargarPlantilla
			// 
			this.btnDescargarPlantilla.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(102)), ((System.Byte)(0)));
			this.btnDescargarPlantilla.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.btnDescargarPlantilla.ForeColor = System.Drawing.Color.White;
			this.btnDescargarPlantilla.Location = new System.Drawing.Point(536, 14);
			this.btnDescargarPlantilla.Name = "btnDescargarPlantilla";
			this.btnDescargarPlantilla.Size = new System.Drawing.Size(160, 23);
			this.btnDescargarPlantilla.TabIndex = 4;
			this.btnDescargarPlantilla.Text = "Descargar Plantilla Excel";
			this.btnDescargarPlantilla.Click += new System.EventHandler(this.btnDescargarPlantilla_Click);
			//
			// btnManual
			//
			this.btnManual.Text      = "? Manual";
			this.btnManual.Location  = new System.Drawing.Point(706, 14);
			this.btnManual.Size      = new System.Drawing.Size(80, 23);
			this.btnManual.TabIndex  = 5;
			this.btnManual.BackColor = System.Drawing.Color.FromArgb(0, 70, 127);
			this.btnManual.ForeColor = System.Drawing.Color.White;
			this.btnManual.Font      = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Bold);
			this.btnManual.Click    += new System.EventHandler(this.btnManual_Click);
			//
			// tabControl
			// 
			this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl.Controls.Add(this.tabCargarExcel);
			this.tabControl.Controls.Add(this.tabPreparar);
			this.tabControl.Controls.Add(this.tabVentaCartera);
			this.tabControl.Location = new System.Drawing.Point(4, 58);
			this.tabControl.Name = "tabControl";
			this.tabControl.SelectedIndex = 0;
			this.tabControl.Size = new System.Drawing.Size(876, 480);
			this.tabControl.TabIndex = 1;
			// 
			// tabCargarExcel
			// 
			this.tabCargarExcel.Controls.Add(this.lblArchivoExcel);
			this.tabCargarExcel.Controls.Add(this.txtArchivoExcel);
			this.tabCargarExcel.Controls.Add(this.btnBuscarExcel);
			this.tabCargarExcel.Controls.Add(this.btnCargarExcel);
			this.tabCargarExcel.Controls.Add(this.lblTotalExcel);
			this.tabCargarExcel.Controls.Add(this.lblBuscarExcel);
			this.tabCargarExcel.Controls.Add(this.txtBuscarExcel);
			this.tabCargarExcel.Controls.Add(this.gridExcel);
			this.tabCargarExcel.Location = new System.Drawing.Point(4, 22);
			this.tabCargarExcel.Name = "tabCargarExcel";
			this.tabCargarExcel.Size = new System.Drawing.Size(868, 454);
			this.tabCargarExcel.TabIndex = 0;
			this.tabCargarExcel.Text = "1. Cargar Excel";
			// 
			// lblArchivoExcel
			// 
			this.lblArchivoExcel.Location = new System.Drawing.Point(6, 14);
			this.lblArchivoExcel.Name = "lblArchivoExcel";
			this.lblArchivoExcel.Size = new System.Drawing.Size(48, 16);
			this.lblArchivoExcel.TabIndex = 0;
			this.lblArchivoExcel.Text = "Archivo:";
			// 
			// txtArchivoExcel
			// 
			this.txtArchivoExcel.Location = new System.Drawing.Point(58, 11);
			this.txtArchivoExcel.Name = "txtArchivoExcel";
			this.txtArchivoExcel.ReadOnly = true;
			this.txtArchivoExcel.Size = new System.Drawing.Size(510, 20);
			this.txtArchivoExcel.TabIndex = 1;
			this.txtArchivoExcel.Text = "";
			// 
			// btnBuscarExcel
			// 
			this.btnBuscarExcel.Location = new System.Drawing.Point(574, 9);
			this.btnBuscarExcel.Name = "btnBuscarExcel";
			this.btnBuscarExcel.Size = new System.Drawing.Size(72, 23);
			this.btnBuscarExcel.TabIndex = 2;
			this.btnBuscarExcel.Text = "Buscar...";
			this.btnBuscarExcel.Click += new System.EventHandler(this.btnBuscarExcel_Click);
			// 
			// btnCargarExcel
			// 
			this.btnCargarExcel.BackColor = System.Drawing.Color.SteelBlue;
			this.btnCargarExcel.ForeColor = System.Drawing.Color.White;
			this.btnCargarExcel.Location = new System.Drawing.Point(652, 9);
			this.btnCargarExcel.Name = "btnCargarExcel";
			this.btnCargarExcel.Size = new System.Drawing.Size(95, 23);
			this.btnCargarExcel.TabIndex = 3;
			this.btnCargarExcel.Text = "Cargar Excel";
			this.btnCargarExcel.Click += new System.EventHandler(this.btnCargarExcel_Click);
			// 
			// lblTotalExcel
			// 
			this.lblTotalExcel.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblTotalExcel.Location = new System.Drawing.Point(6, 38);
			this.lblTotalExcel.Name = "lblTotalExcel";
			this.lblTotalExcel.Size = new System.Drawing.Size(200, 16);
			this.lblTotalExcel.TabIndex = 4;
			this.lblTotalExcel.Text = "Total leidos: 0";
			// 
			// gridExcel
			// 
			this.gridExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridExcel.CaptionText = "Datos leidos del Excel";
			this.gridExcel.DataMember = "";
			this.gridExcel.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.gridExcel.Location = new System.Drawing.Point(4, 58);
			this.gridExcel.Name = "gridExcel";
			this.gridExcel.ReadOnly = true;
			this.gridExcel.Size = new System.Drawing.Size(1520, 740);
			this.gridExcel.TabIndex = 5;
			// 
			// tabPreparar
			// 
			this.tabPreparar.Controls.Add(this.btnPreparar);
			this.tabPreparar.Controls.Add(this.lblContTotal);
			this.tabPreparar.Controls.Add(this.lblContEncontrados);
			this.tabPreparar.Controls.Add(this.lblContSinCoincidencia);
			this.tabPreparar.Controls.Add(this.lblContYaProcesados);
			this.tabPreparar.Controls.Add(this.tabValidacion);
			this.tabPreparar.Location = new System.Drawing.Point(4, 22);
			this.tabPreparar.Name = "tabPreparar";
			this.tabPreparar.Size = new System.Drawing.Size(868, 454);
			this.tabPreparar.TabIndex = 1;
			this.tabPreparar.Text = "2. Preparar y Validar";
			// 
			// btnPreparar
			// 
			this.btnPreparar.BackColor = System.Drawing.Color.DarkGreen;
			this.btnPreparar.Enabled = false;
			this.btnPreparar.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.btnPreparar.ForeColor = System.Drawing.Color.White;
			this.btnPreparar.Location = new System.Drawing.Point(6, 9);
			this.btnPreparar.Name = "btnPreparar";
			this.btnPreparar.Size = new System.Drawing.Size(120, 26);
			this.btnPreparar.TabIndex = 0;
			this.btnPreparar.Text = "Preparar Datos";
			this.btnPreparar.Click += new System.EventHandler(this.btnPreparar_Click);
			// 
			// lblContTotal
			// 
			this.lblContTotal.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblContTotal.Location = new System.Drawing.Point(136, 14);
			this.lblContTotal.Name = "lblContTotal";
			this.lblContTotal.Size = new System.Drawing.Size(130, 16);
			this.lblContTotal.TabIndex = 1;
			this.lblContTotal.Text = "Total Excel: 0";
			// 
			// lblContEncontrados
			// 
			this.lblContEncontrados.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblContEncontrados.ForeColor = System.Drawing.Color.Green;
			this.lblContEncontrados.Location = new System.Drawing.Point(278, 14);
			this.lblContEncontrados.Name = "lblContEncontrados";
			this.lblContEncontrados.Size = new System.Drawing.Size(140, 16);
			this.lblContEncontrados.TabIndex = 2;
			this.lblContEncontrados.Text = "Encontrados: 0";
			//
			// lblContSinCoincidencia
			//
			this.lblContSinCoincidencia.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblContSinCoincidencia.ForeColor = System.Drawing.Color.OrangeRed;
			this.lblContSinCoincidencia.Location = new System.Drawing.Point(430, 14);
			this.lblContSinCoincidencia.Name = "lblContSinCoincidencia";
			this.lblContSinCoincidencia.Size = new System.Drawing.Size(170, 16);
			this.lblContSinCoincidencia.TabIndex = 3;
			this.lblContSinCoincidencia.Text = "Sin coincidencia: 0";
			//
			// lblContYaProcesados
			//
			this.lblContYaProcesados.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblContYaProcesados.ForeColor = System.Drawing.Color.Green;
			this.lblContYaProcesados.Location = new System.Drawing.Point(614, 14);
			this.lblContYaProcesados.Name = "lblContYaProcesados";
			this.lblContYaProcesados.Size = new System.Drawing.Size(170, 16);
			this.lblContYaProcesados.TabIndex = 5;
			this.lblContYaProcesados.Text = "Ya procesados: 0";
			// 
			// tabValidacion
			// 
			this.tabValidacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabValidacion.Controls.Add(this.tabEncontrados);
			this.tabValidacion.Controls.Add(this.tabSinCoincidencia);
			this.tabValidacion.Location = new System.Drawing.Point(4, 42);
			this.tabValidacion.Name = "tabValidacion";
			this.tabValidacion.SelectedIndex = 0;
			this.tabValidacion.Size = new System.Drawing.Size(1520, 756);
			this.tabValidacion.TabIndex = 4;
			// 
			// tabEncontrados
			// 
			this.tabEncontrados.Controls.Add(this.lblBuscarEncontrados);
			this.tabEncontrados.Controls.Add(this.txtBuscarEncontrados);
			this.tabEncontrados.Controls.Add(this.gridEncontrados);
			this.tabEncontrados.Location = new System.Drawing.Point(4, 22);
			this.tabEncontrados.Name = "tabEncontrados";
			this.tabEncontrados.Size = new System.Drawing.Size(1512, 730);
			this.tabEncontrados.TabIndex = 0;
			this.tabEncontrados.Text = "ENCONTRADOS";
			// 
			// gridEncontrados
			// 
			this.gridEncontrados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right)));
			this.gridEncontrados.CaptionText = "Registros encontrados en Compra";
			this.gridEncontrados.DataMember = "";
			this.gridEncontrados.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.gridEncontrados.Location = new System.Drawing.Point(0, 30);
			this.gridEncontrados.Name = "gridEncontrados";
			this.gridEncontrados.ReadOnly = true;
			this.gridEncontrados.Size = new System.Drawing.Size(1512, 700);
			this.gridEncontrados.TabIndex = 3;
			// 
			// tabSinCoincidencia
			// 
			this.tabSinCoincidencia.Controls.Add(this.lblBuscarSinCoinc);
			this.tabSinCoincidencia.Controls.Add(this.txtBuscarSinCoinc);
			this.tabSinCoincidencia.Controls.Add(this.gridSinCoincidencia);
			this.tabSinCoincidencia.Location = new System.Drawing.Point(4, 22);
			this.tabSinCoincidencia.Name = "tabSinCoincidencia";
			this.tabSinCoincidencia.Size = new System.Drawing.Size(1512, 730);
			this.tabSinCoincidencia.TabIndex = 1;
			this.tabSinCoincidencia.Text = "SIN COINCIDENCIA";
			// 
			// gridSinCoincidencia
			// 
			this.gridSinCoincidencia.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right)));
			this.gridSinCoincidencia.CaptionText = "Registros sin coincidencia en Compra";
			this.gridSinCoincidencia.DataMember = "";
			this.gridSinCoincidencia.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.gridSinCoincidencia.Location = new System.Drawing.Point(0, 30);
			this.gridSinCoincidencia.Name = "gridSinCoincidencia";
			this.gridSinCoincidencia.ReadOnly = true;
			this.gridSinCoincidencia.Size = new System.Drawing.Size(1512, 700);
			this.gridSinCoincidencia.TabIndex = 3;
			// 
			// tabVentaCartera
			// 
			this.tabVentaCartera.Controls.Add(this.btnEjecutarVenta);
			this.tabVentaCartera.Controls.Add(this.lblResumenTotal);
			this.tabVentaCartera.Controls.Add(this.lblResumenActualizados);
			this.tabVentaCartera.Controls.Add(this.lblResumenFallidos);
			this.tabVentaCartera.Controls.Add(this.tabResultadoVenta);
			this.tabVentaCartera.Location = new System.Drawing.Point(4, 22);
			this.tabVentaCartera.Name = "tabVentaCartera";
			this.tabVentaCartera.Size = new System.Drawing.Size(868, 454);
			this.tabVentaCartera.TabIndex = 2;
			this.tabVentaCartera.Text = "3. Ejecutar Venta";
			// 
			// btnEjecutarVenta
			// 
			this.btnEjecutarVenta.BackColor = System.Drawing.Color.DarkRed;
			this.btnEjecutarVenta.Enabled = false;
			this.btnEjecutarVenta.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
			this.btnEjecutarVenta.ForeColor = System.Drawing.Color.White;
			this.btnEjecutarVenta.Location = new System.Drawing.Point(744, 8);
			this.btnEjecutarVenta.Name = "btnEjecutarVenta";
			this.btnEjecutarVenta.Size = new System.Drawing.Size(110, 40);
			this.btnEjecutarVenta.TabIndex = 2;
			this.btnEjecutarVenta.Text = "Ejecutar Venta";
			this.btnEjecutarVenta.Click += new System.EventHandler(this.btnEjecutarVenta_Click);
			// 
			// lblResumenTotal
			// 
			this.lblResumenTotal.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblResumenTotal.Location = new System.Drawing.Point(8, 16);
			this.lblResumenTotal.Name = "lblResumenTotal";
			this.lblResumenTotal.Size = new System.Drawing.Size(100, 16);
			this.lblResumenTotal.TabIndex = 3;
			this.lblResumenTotal.Text = "Total: 0";
			// 
			// lblResumenActualizados
			// 
			this.lblResumenActualizados.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblResumenActualizados.ForeColor = System.Drawing.Color.Green;
			this.lblResumenActualizados.Location = new System.Drawing.Point(112, 16);
			this.lblResumenActualizados.Name = "lblResumenActualizados";
			this.lblResumenActualizados.Size = new System.Drawing.Size(150, 16);
			this.lblResumenActualizados.TabIndex = 4;
			this.lblResumenActualizados.Text = "Actualizados: 0";
			// 
			// lblResumenFallidos
			// 
			this.lblResumenFallidos.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblResumenFallidos.ForeColor = System.Drawing.Color.OrangeRed;
			this.lblResumenFallidos.Location = new System.Drawing.Point(272, 16);
			this.lblResumenFallidos.Name = "lblResumenFallidos";
			this.lblResumenFallidos.Size = new System.Drawing.Size(170, 16);
			this.lblResumenFallidos.TabIndex = 5;
			this.lblResumenFallidos.Text = "Sin coincidencia: 0";
			// 
			// tabResultadoVenta
			// 
			this.tabResultadoVenta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabResultadoVenta.Controls.Add(this.tabActualizados);
			this.tabResultadoVenta.Controls.Add(this.tabFallidos);
			this.tabResultadoVenta.Location = new System.Drawing.Point(4, 56);
			this.tabResultadoVenta.Name = "tabResultadoVenta";
			this.tabResultadoVenta.SelectedIndex = 0;
			this.tabResultadoVenta.Size = new System.Drawing.Size(1520, 736);
			this.tabResultadoVenta.TabIndex = 6;
			// 
			// tabActualizados
			// 
			this.tabActualizados.Controls.Add(this.lblBuscarActualizados);
			this.tabActualizados.Controls.Add(this.txtBuscarActualizados);
			this.tabActualizados.Controls.Add(this.gridActualizados);
			this.tabActualizados.Location = new System.Drawing.Point(4, 22);
			this.tabActualizados.Name = "tabActualizados";
			this.tabActualizados.Size = new System.Drawing.Size(1512, 710);
			this.tabActualizados.TabIndex = 0;
			this.tabActualizados.Text = "ACTUALIZADOS";
			// 
			// gridActualizados
			// 
			this.gridActualizados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right)));
			this.gridActualizados.CaptionText = "Registros Actualizados en Compra";
			this.gridActualizados.DataMember = "";
			this.gridActualizados.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.gridActualizados.Location = new System.Drawing.Point(0, 30);
			this.gridActualizados.Name = "gridActualizados";
			this.gridActualizados.ReadOnly = true;
			this.gridActualizados.Size = new System.Drawing.Size(1512, 680);
			this.gridActualizados.TabIndex = 3;
			// 
			// tabFallidos
			// 
			this.tabFallidos.Controls.Add(this.lblBuscarFallidos);
			this.tabFallidos.Controls.Add(this.txtBuscarFallidos);
			this.tabFallidos.Controls.Add(this.gridFallidos);
			this.tabFallidos.Location = new System.Drawing.Point(4, 22);
			this.tabFallidos.Name = "tabFallidos";
			this.tabFallidos.Size = new System.Drawing.Size(1512, 694);
			this.tabFallidos.TabIndex = 1;
			this.tabFallidos.Text = "SIN COINCIDENCIA";
			// 
			// gridFallidos
			// 
			this.gridFallidos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right)));
			this.gridFallidos.CaptionText = "Registros sin coincidencia";
			this.gridFallidos.DataMember = "";
			this.gridFallidos.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.gridFallidos.Location = new System.Drawing.Point(0, 30);
			this.gridFallidos.Name = "gridFallidos";
			this.gridFallidos.ReadOnly = true;
			this.gridFallidos.Size = new System.Drawing.Size(1512, 664);
			this.gridFallidos.TabIndex = 3;
			// 
			// lblBuscarExcel
			// 
			this.lblBuscarExcel.Location = new System.Drawing.Point(260, 38);
			this.lblBuscarExcel.Name = "lblBuscarExcel";
			this.lblBuscarExcel.Size = new System.Drawing.Size(46, 16);
			this.lblBuscarExcel.TabIndex = 5;
			this.lblBuscarExcel.Text = "Buscar:";
			// 
			// txtBuscarExcel
			// 
			this.txtBuscarExcel.Location = new System.Drawing.Point(310, 35);
			this.txtBuscarExcel.Name = "txtBuscarExcel";
			this.txtBuscarExcel.Size = new System.Drawing.Size(220, 20);
			this.txtBuscarExcel.TabIndex = 6;
			this.txtBuscarExcel.Text = "";
			this.txtBuscarExcel.TextChanged += new System.EventHandler(this.txtBuscarExcel_TextChanged);
			// 
			// lblBuscarEncontrados
			// 
			this.lblBuscarEncontrados.Location = new System.Drawing.Point(4, 6);
			this.lblBuscarEncontrados.Name = "lblBuscarEncontrados";
			this.lblBuscarEncontrados.Size = new System.Drawing.Size(46, 16);
			this.lblBuscarEncontrados.TabIndex = 1;
			this.lblBuscarEncontrados.Text = "Buscar:";
			//
			// txtBuscarEncontrados
			//
			this.txtBuscarEncontrados.Location = new System.Drawing.Point(54, 4);
			this.txtBuscarEncontrados.Name = "txtBuscarEncontrados";
			this.txtBuscarEncontrados.Size = new System.Drawing.Size(220, 20);
			this.txtBuscarEncontrados.TabIndex = 2;
			this.txtBuscarEncontrados.Text = "";
			this.txtBuscarEncontrados.TextChanged += new System.EventHandler(this.txtBuscarEncontrados_TextChanged);
			//
			// lblBuscarSinCoinc
			//
			this.lblBuscarSinCoinc.Location = new System.Drawing.Point(4, 6);
			this.lblBuscarSinCoinc.Name = "lblBuscarSinCoinc";
			this.lblBuscarSinCoinc.Size = new System.Drawing.Size(46, 16);
			this.lblBuscarSinCoinc.TabIndex = 1;
			this.lblBuscarSinCoinc.Text = "Buscar:";
			//
			// txtBuscarSinCoinc
			//
			this.txtBuscarSinCoinc.Location = new System.Drawing.Point(54, 4);
			this.txtBuscarSinCoinc.Name = "txtBuscarSinCoinc";
			this.txtBuscarSinCoinc.Size = new System.Drawing.Size(220, 20);
			this.txtBuscarSinCoinc.TabIndex = 2;
			this.txtBuscarSinCoinc.Text = "";
			this.txtBuscarSinCoinc.TextChanged += new System.EventHandler(this.txtBuscarSinCoinc_TextChanged);
			//
			// lblBuscarActualizados
			//
			this.lblBuscarActualizados.Location = new System.Drawing.Point(4, 6);
			this.lblBuscarActualizados.Name = "lblBuscarActualizados";
			this.lblBuscarActualizados.Size = new System.Drawing.Size(46, 16);
			this.lblBuscarActualizados.TabIndex = 1;
			this.lblBuscarActualizados.Text = "Buscar:";
			//
			// txtBuscarActualizados
			//
			this.txtBuscarActualizados.Location = new System.Drawing.Point(54, 4);
			this.txtBuscarActualizados.Name = "txtBuscarActualizados";
			this.txtBuscarActualizados.Size = new System.Drawing.Size(220, 20);
			this.txtBuscarActualizados.TabIndex = 2;
			this.txtBuscarActualizados.Text = "";
			this.txtBuscarActualizados.TextChanged += new System.EventHandler(this.txtBuscarActualizados_TextChanged);
			//
			// lblBuscarFallidos
			//
			this.lblBuscarFallidos.Location = new System.Drawing.Point(4, 6);
			this.lblBuscarFallidos.Name = "lblBuscarFallidos";
			this.lblBuscarFallidos.Size = new System.Drawing.Size(46, 16);
			this.lblBuscarFallidos.TabIndex = 1;
			this.lblBuscarFallidos.Text = "Buscar:";
			//
			// txtBuscarFallidos
			//
			this.txtBuscarFallidos.Location = new System.Drawing.Point(54, 4);
			this.txtBuscarFallidos.Name = "txtBuscarFallidos";
			this.txtBuscarFallidos.Size = new System.Drawing.Size(220, 20);
			this.txtBuscarFallidos.TabIndex = 2;
			this.txtBuscarFallidos.Text = "";
			this.txtBuscarFallidos.TextChanged += new System.EventHandler(this.txtBuscarFallidos_TextChanged);
			// 
			// statusBar
			// 
			this.statusBar.Location = new System.Drawing.Point(0, 558);
			this.statusBar.Name = "statusBar";
			this.statusBar.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
																																								 this.sbpMensaje});
			this.statusBar.ShowPanels = true;
			this.statusBar.Size = new System.Drawing.Size(884, 22);
			this.statusBar.TabIndex = 3;
			// 
			// sbpMensaje
			// 
			this.sbpMensaje.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring;
			this.sbpMensaje.Text = "Listo";
			this.sbpMensaje.Width = 868;
			// 
			// progressBar
			// 
			this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.progressBar.Location = new System.Drawing.Point(4, 542);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(876, 14);
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
			// frmFideicomisoProceso
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(884, 580);
			this.Controls.Add(this.grpProceso);
			this.Controls.Add(this.tabControl);
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.statusBar);
			this.MinimumSize = new System.Drawing.Size(800, 600);
			this.Name = "frmFideicomisoProceso";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Venta de Cartera Fideicomiso";
			this.Load += new System.EventHandler(this.frmFideicomisoProceso_Load);
			this.grpProceso.ResumeLayout(false);
			this.tabControl.ResumeLayout(false);
			this.tabCargarExcel.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridExcel)).EndInit();
			this.tabPreparar.ResumeLayout(false);
			this.tabValidacion.ResumeLayout(false);
			this.tabEncontrados.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridEncontrados)).EndInit();
			this.tabSinCoincidencia.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridSinCoincidencia)).EndInit();
			this.tabVentaCartera.ResumeLayout(false);
			this.tabResultadoVenta.ResumeLayout(false);
			this.tabActualizados.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridActualizados)).EndInit();
			this.tabFallidos.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridFallidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion
	}
}
