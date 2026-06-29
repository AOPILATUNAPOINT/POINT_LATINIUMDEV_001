using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	public class frmCierreDescuentoSegmentoConfig : System.Windows.Forms.Form
	{
		// --- Grid maestro ---
		private System.Windows.Forms.DataGrid   gridSegmentos;
		private System.Windows.Forms.Label      lblListaSegmentos;

		// --- Panel detalle ---
		private System.Windows.Forms.GroupBox   grpDetalle;
		private System.Windows.Forms.Label      lblSegmentoNombre;

		// Cuenta Cartera
		private System.Windows.Forms.Label      lblCartera;
		private System.Windows.Forms.TextBox    txtCarteraCodigo;
		private System.Windows.Forms.TextBox    txtCarteraNombre;
		private System.Windows.Forms.Button     btnBuscarCartera;
		private System.Windows.Forms.Label      lblCarteraHint;

		// Cuenta Descuento/Castigo
		private System.Windows.Forms.Label      lblDescuento;
		private System.Windows.Forms.TextBox    txtDescuentoCodigo;
		private System.Windows.Forms.TextBox    txtDescuentoNombre;
		private System.Windows.Forms.Button     btnBuscarDescuento;
		private System.Windows.Forms.Label      lblDescuentoHint;

		// Cuenta Recaudo
		private System.Windows.Forms.Label      lblRecaudo;
		private System.Windows.Forms.TextBox    txtRecaudoCodigo;
		private System.Windows.Forms.TextBox    txtRecaudoNombre;
		private System.Windows.Forms.Button     btnBuscarRecaudo;
		private System.Windows.Forms.Label      lblRecaudoHint;

		// Referencia hint
		private System.Windows.Forms.Label      lblRefs;

		// Activo
		private System.Windows.Forms.CheckBox   chkActivo;

		// Botones detalle
		private System.Windows.Forms.Button     btnGuardar;
		private System.Windows.Forms.Button     btnCancelar;

		// --- Barra inferior ---
		private System.Windows.Forms.StatusBar      statusBar;
		private System.Windows.Forms.StatusBarPanel sbpMensaje;

		// --- Estado ---
		private C1.Data.C1DataSet cdsSeteoF;
		private Acceso            miAcceso;
		private DataTable         dtSegmentos;
		private int               idSegmentoActual = 0;
		private bool              modoEdicion      = false;

		private System.ComponentModel.Container components = null;

		public frmCierreDescuentoSegmentoConfig()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null) components.Dispose();
			base.Dispose(disposing);
		}

		// ================================================================
		//  CARGA
		// ================================================================

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCierreDescuentoSegmentoConfig_Load(object sender, System.EventArgs e)
		{
			try
			{
				miAcceso  = new Acceso(cdsSeteoF, "09100002");

				if (!Funcion.Permiso("9101", cdsSeteoF))
				{
					MessageBox.Show("No tiene acceso a la configuración de Cierre con Descuento.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.Close();
					return;
				}

				CargarSegmentos();
				LimpiarDetalle();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al iniciar:\n" + ex.Message, "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private SqlConnection AbrirConexion()
		{
			SqlConnection cn = new SqlConnection(
				"Password=Bl45o6$9;Persist Security Info=True;User ID=" + MenuLatinium.stUsuario.Trim()
				+ ";Initial Catalog=" + MenuLatinium.stDirFacturacion.Trim()
				+ ";Data Source=" + MenuLatinium.stDirServidor.Trim()
				+ ";Max Pool Size=1024;");
			cn.Open();
			return cn;
		}

		private void SetStatus(string msg) { sbpMensaje.Text = msg; }

		// ================================================================
		//  CARGAR GRID MAESTRO
		// ================================================================

		private void CargarSegmentos()
		{
			SqlConnection cn = null;
			try
			{
				cn = AbrirConexion();
				SqlDataAdapter da = new SqlDataAdapter(
					"SELECT s.idCierreDescuentoSegmento, s.Segmento, " +
					"       s.idCuentaCartera,   cc.Codigo AS CodigoCartera,   cc.Descripcion AS NombreCartera, " +
					"       s.idCuentaDescuento, cd.Codigo AS CodigoDescuento, cd.Descripcion AS NombreDescuento, " +
					"       s.idCuentaRecaudo,   cr.Codigo AS CodigoRecaudo,   cr.Descripcion AS NombreRecaudo, " +
					"       s.Activo " +
					"FROM CierreDescuentoSegmento s " +
					"LEFT JOIN Cuenta cc ON cc.idCuenta = s.idCuentaCartera   AND s.idCuentaCartera   <> 999999 " +
					"LEFT JOIN Cuenta cd ON cd.idCuenta = s.idCuentaDescuento AND s.idCuentaDescuento <> 999999 " +
					"LEFT JOIN Cuenta cr ON cr.idCuenta = s.idCuentaRecaudo   AND s.idCuentaRecaudo   <> 999999 " +
					"ORDER BY s.Segmento", cn);
				dtSegmentos = new DataTable();
				da.Fill(dtSegmentos);

				gridSegmentos.DataSource = dtSegmentos;

				// Ocultar columnas de ID en el grid
				DataGridTableStyle ts = new DataGridTableStyle();
				ts.MappingName = dtSegmentos.TableName;
				AgregarColumnaGrid(ts, "Segmento",         "Segmento",          180, false);
				AgregarColumnaGrid(ts, "CodigoCartera",    "Cta. Cartera",       90, true);
				AgregarColumnaGrid(ts, "NombreCartera",    "Nombre Cartera",    180, true);
				AgregarColumnaGrid(ts, "CodigoDescuento",  "Cta. Descuento",     90, true);
				AgregarColumnaGrid(ts, "NombreDescuento",  "Nombre Descuento",  180, true);
				AgregarColumnaGrid(ts, "CodigoRecaudo",    "Cta. Recaudo",       90, true);
				AgregarColumnaGrid(ts, "NombreRecaudo",    "Nombre Recaudo",    180, true);
				AgregarColumnaGrid(ts, "Activo",           "Activo",             50, true);
				gridSegmentos.TableStyles.Clear();
				gridSegmentos.TableStyles.Add(ts);

				ActualizarEstadoConfiguracion();
			}
			finally
			{
				if (cn != null) cn.Close();
			}
		}

		private void AgregarColumnaGrid(DataGridTableStyle ts, string campo, string encabezado, int ancho, bool soloLectura)
		{
			DataGridTextBoxColumn col = new DataGridTextBoxColumn();
			col.MappingName  = campo;
			col.HeaderText   = encabezado;
			col.Width        = ancho;
			col.ReadOnly     = soloLectura;
			ts.GridColumnStyles.Add(col);
		}

		private void ActualizarEstadoConfiguracion()
		{
			if (dtSegmentos == null) return;
			int sinConfig = 0;
			foreach (DataRow r in dtSegmentos.Rows)
			{
				if (r["CodigoCartera"]   == DBNull.Value ||
					r["CodigoDescuento"] == DBNull.Value ||
					r["CodigoRecaudo"]   == DBNull.Value)
					sinConfig++;
			}
			if (sinConfig > 0)
				SetStatus("ATENCIÓN: " + sinConfig + " segmento(s) sin cuentas contables configuradas.");
			else
				SetStatus("Todos los segmentos están configurados correctamente.");
		}

		// ================================================================
		//  SELECCIÓN EN GRID
		// ================================================================

		private void gridSegmentos_Click(object sender, System.EventArgs e)
		{
			CargarDetalle();
		}

		private void gridSegmentos_CurrentCellChanged(object sender, System.EventArgs e)
		{
			CargarDetalle();
		}

		private void CargarDetalle()
		{
			if (dtSegmentos == null || dtSegmentos.Rows.Count == 0) return;
			int fila = gridSegmentos.CurrentRowIndex;
			if (fila < 0 || fila >= dtSegmentos.Rows.Count) return;

			DataRow r = dtSegmentos.Rows[fila];
			idSegmentoActual = Convert.ToInt32(r["idCierreDescuentoSegmento"]);

			lblSegmentoNombre.Text = "Segmento: " + r["Segmento"].ToString().ToUpper();

			// Cartera
			txtCarteraCodigo.Text  = r["CodigoCartera"]   == DBNull.Value ? "" : r["CodigoCartera"].ToString();
			txtCarteraNombre.Text  = r["NombreCartera"]   == DBNull.Value ? "" : r["NombreCartera"].ToString();

			// Descuento
			txtDescuentoCodigo.Text = r["CodigoDescuento"] == DBNull.Value ? "" : r["CodigoDescuento"].ToString();
			txtDescuentoNombre.Text = r["NombreDescuento"] == DBNull.Value ? "" : r["NombreDescuento"].ToString();

			// Recaudo
			txtRecaudoCodigo.Text  = r["CodigoRecaudo"]   == DBNull.Value ? "" : r["CodigoRecaudo"].ToString();
			txtRecaudoNombre.Text  = r["NombreRecaudo"]   == DBNull.Value ? "" : r["NombreRecaudo"].ToString();

			chkActivo.Checked = r["Activo"] != DBNull.Value && Convert.ToBoolean(r["Activo"]);

			ColorearEstado();
			grpDetalle.Enabled = true;
			modoEdicion        = false;
		}

		private void ColorearEstado()
		{
			txtCarteraCodigo.BackColor  = txtCarteraCodigo.Text  == "" ? Color.MistyRose : Color.White;
			txtDescuentoCodigo.BackColor = txtDescuentoCodigo.Text == "" ? Color.MistyRose : Color.White;
			txtRecaudoCodigo.BackColor  = txtRecaudoCodigo.Text  == "" ? Color.MistyRose : Color.White;
		}

		private void LimpiarDetalle()
		{
			idSegmentoActual        = 0;
			lblSegmentoNombre.Text  = "Seleccione un segmento de la lista";
			txtCarteraCodigo.Text   = "";
			txtCarteraNombre.Text   = "";
			txtDescuentoCodigo.Text = "";
			txtDescuentoNombre.Text = "";
			txtRecaudoCodigo.Text   = "";
			txtRecaudoNombre.Text   = "";
			chkActivo.Checked       = true;
			grpDetalle.Enabled      = false;
			modoEdicion             = false;
		}

		// ================================================================
		//  BÚSQUEDA DE CUENTAS (F3 / botón)
		// ================================================================

		private void BuscarCuenta(TextBox txtCodigo, TextBox txtNombre)
		{
			string codigo = txtCodigo.Text.Trim();
			if (codigo != "")
			{
				// Buscar por código directo
				SqlConnection cn = null;
				try
				{
					cn = AbrirConexion();
					SqlCommand cmd = new SqlCommand(
						"SELECT TOP 1 idCuenta, Codigo, Descripcion FROM Cuenta " +
						"WHERE Codigo = @Cod AND Activo = 1", cn);
					cmd.Parameters.Add("@Cod", SqlDbType.VarChar).Value = codigo;
					SqlDataReader dr = cmd.ExecuteReader();
					if (dr.Read())
					{
						txtCodigo.Tag  = dr["idCuenta"];
						txtNombre.Text = dr["Descripcion"].ToString();
						dr.Close();
						modoEdicion = true;
						ColorearEstado();
						return;
					}
					dr.Close();
				}
				finally { if (cn != null) cn.Close(); }
			}

			// Abrir buscador de cuentas
			using (BuscaCuentas busqueda = new BuscaCuentas(codigo, 1))
			{
				if (busqueda.ShowDialog() == DialogResult.OK)
				{
					txtCodigo.Text = busqueda.txtCodigo.Value == null ? "" : busqueda.txtCodigo.Value.ToString();
					txtCodigo.Tag  = busqueda.txtIdBusca.Value;
					txtNombre.Text = busqueda.txtDescripcion.Value == null ? "" : busqueda.txtDescripcion.Value.ToString();
					modoEdicion    = true;
					ColorearEstado();
				}
			}
		}

		private void txtCarteraCodigo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3 || e.KeyCode == Keys.Enter)
				BuscarCuenta(txtCarteraCodigo, txtCarteraNombre);
		}

		private void txtDescuentoCodigo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3 || e.KeyCode == Keys.Enter)
				BuscarCuenta(txtDescuentoCodigo, txtDescuentoNombre);
		}

		private void txtRecaudoCodigo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3 || e.KeyCode == Keys.Enter)
				BuscarCuenta(txtRecaudoCodigo, txtRecaudoNombre);
		}

		private void btnBuscarCartera_Click(object sender, System.EventArgs e)
		{
			BuscarCuenta(txtCarteraCodigo, txtCarteraNombre);
		}

		private void btnBuscarDescuento_Click(object sender, System.EventArgs e)
		{
			BuscarCuenta(txtDescuentoCodigo, txtDescuentoNombre);
		}

		private void btnBuscarRecaudo_Click(object sender, System.EventArgs e)
		{
			BuscarCuenta(txtRecaudoCodigo, txtRecaudoNombre);
		}

		// ================================================================
		//  GUARDAR
		// ================================================================

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (idSegmentoActual == 0)
			{
				MessageBox.Show("Seleccione un segmento de la lista.", "Aviso",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Obtener idCuenta de cada campo
			int idCartera   = ObtenerIdCuenta(txtCarteraCodigo,   txtCarteraNombre,   "Cuenta Cartera");
			if (idCartera   == -1) return;
			int idDescuento = ObtenerIdCuenta(txtDescuentoCodigo, txtDescuentoNombre, "Cuenta Descuento");
			if (idDescuento == -1) return;
			int idRecaudo   = ObtenerIdCuenta(txtRecaudoCodigo,   txtRecaudoNombre,   "Cuenta Recaudo");
			if (idRecaudo   == -1) return;

			// Confirmar
			if (MessageBox.Show(
				"¿Confirma guardar la configuración del segmento?\n\n" +
				lblSegmentoNombre.Text + "\n\n" +
				"Cartera:   " + txtCarteraCodigo.Text + " - " + txtCarteraNombre.Text + "\n" +
				"Descuento: " + txtDescuentoCodigo.Text + " - " + txtDescuentoNombre.Text + "\n" +
				"Recaudo:   " + txtRecaudoCodigo.Text + " - " + txtRecaudoNombre.Text,
				"Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
				return;

			SqlConnection cn = null;
			try
			{
				cn = AbrirConexion();
				SqlCommand cmd = new SqlCommand(
					"UPDATE CierreDescuentoSegmento " +
					"SET idCuentaCartera   = @idCartera, " +
					"    idCuentaDescuento = @idDescuento, " +
					"    idCuentaRecaudo   = @idRecaudo, " +
					"    Activo            = @Activo " +
					"WHERE idCierreDescuentoSegmento = @id", cn);
				cmd.Parameters.Add("@idCartera",   SqlDbType.Int).Value = idCartera;
				cmd.Parameters.Add("@idDescuento", SqlDbType.Int).Value = idDescuento;
				cmd.Parameters.Add("@idRecaudo",   SqlDbType.Int).Value = idRecaudo;
				cmd.Parameters.Add("@Activo",      SqlDbType.Bit).Value = chkActivo.Checked ? 1 : 0;
				cmd.Parameters.Add("@id",          SqlDbType.Int).Value = idSegmentoActual;
				cmd.ExecuteNonQuery();

				SetStatus("Segmento guardado correctamente: " + lblSegmentoNombre.Text);
				modoEdicion = false;
				CargarSegmentos();

				// Reseleccionar la fila guardada
				for (int i = 0; i < dtSegmentos.Rows.Count; i++)
				{
					if (Convert.ToInt32(dtSegmentos.Rows[i]["idCierreDescuentoSegmento"]) == idSegmentoActual)
					{
						gridSegmentos.CurrentRowIndex = i;
						break;
					}
				}
				CargarDetalle();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al guardar:\n" + ex.Message, "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally { if (cn != null) cn.Close(); }
		}

		private int ObtenerIdCuenta(TextBox txtCodigo, TextBox txtNombre, string etiqueta)
		{
			// Si ya tiene Tag con el idCuenta (viene del buscador) usarlo
			if (txtCodigo.Tag != null && txtCodigo.Tag.ToString() != "")
			{
				int id = 0;
				try { id = Convert.ToInt32(txtCodigo.Tag); } catch { id = 0; }
				if (id > 0) return id;
			}

			// Si no, buscar por código en BD
			string codigo = txtCodigo.Text.Trim();
			if (codigo == "")
			{
				MessageBox.Show("La " + etiqueta + " es obligatoria.\nUse F3 o el botón [...] para seleccionar.",
					"Campo requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				txtCodigo.Focus();
				return -1;
			}

			SqlConnection cn = null;
			try
			{
				cn = AbrirConexion();
				SqlCommand cmd = new SqlCommand(
					"SELECT TOP 1 idCuenta, Descripcion FROM Cuenta WHERE Codigo = @Cod AND Activo = 1", cn);
				cmd.Parameters.Add("@Cod", SqlDbType.VarChar).Value = codigo;
				SqlDataReader dr = cmd.ExecuteReader();
				if (dr.Read())
				{
					int idCuenta = Convert.ToInt32(dr["idCuenta"]);
					txtNombre.Text = dr["Descripcion"].ToString();
					txtCodigo.Tag  = idCuenta;
					dr.Close();
					return idCuenta;
				}
				dr.Close();
			}
			finally { if (cn != null) cn.Close(); }

			MessageBox.Show("Código de cuenta no encontrado: " + codigo + "\nVerifique e intente nuevamente.",
				etiqueta + " no encontrada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			txtCodigo.Focus();
			return -1;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (modoEdicion &&
				MessageBox.Show("¿Descartar los cambios?", "Confirmar",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
				return;

			CargarDetalle();
		}

		// ================================================================
		//  INITIALIZE COMPONENT
		// ================================================================

		private void InitializeComponent()
		{
			this.gridSegmentos        = new System.Windows.Forms.DataGrid();
			this.lblListaSegmentos    = new System.Windows.Forms.Label();
			this.grpDetalle           = new System.Windows.Forms.GroupBox();
			this.lblSegmentoNombre    = new System.Windows.Forms.Label();

			this.lblCartera           = new System.Windows.Forms.Label();
			this.txtCarteraCodigo     = new System.Windows.Forms.TextBox();
			this.txtCarteraNombre     = new System.Windows.Forms.TextBox();
			this.btnBuscarCartera     = new System.Windows.Forms.Button();
			this.lblCarteraHint       = new System.Windows.Forms.Label();

			this.lblDescuento         = new System.Windows.Forms.Label();
			this.txtDescuentoCodigo   = new System.Windows.Forms.TextBox();
			this.txtDescuentoNombre   = new System.Windows.Forms.TextBox();
			this.btnBuscarDescuento   = new System.Windows.Forms.Button();
			this.lblDescuentoHint     = new System.Windows.Forms.Label();

			this.lblRecaudo           = new System.Windows.Forms.Label();
			this.txtRecaudoCodigo     = new System.Windows.Forms.TextBox();
			this.txtRecaudoNombre     = new System.Windows.Forms.TextBox();
			this.btnBuscarRecaudo     = new System.Windows.Forms.Button();
			this.lblRecaudoHint       = new System.Windows.Forms.Label();

			this.chkActivo            = new System.Windows.Forms.CheckBox();
			this.btnGuardar           = new System.Windows.Forms.Button();
			this.btnCancelar          = new System.Windows.Forms.Button();

			this.statusBar            = new System.Windows.Forms.StatusBar();
			this.sbpMensaje           = new System.Windows.Forms.StatusBarPanel();
			this.cdsSeteoF            = new C1.Data.C1DataSet();

			((System.ComponentModel.ISupportInitialize)(this.gridSegmentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();

			// Form
			this.Text          = "Configuración — Cuentas Contables por Segmento (Cierre con Descuento)";
			this.Size          = new Size(950, 600);
			this.StartPosition = FormStartPosition.CenterScreen;
			this.Load         += new EventHandler(this.frmCierreDescuentoSegmentoConfig_Load);

			// lblListaSegmentos
			this.lblListaSegmentos.Text      = "Segmentos configurados — seleccione uno para editar sus cuentas:";
			this.lblListaSegmentos.Location  = new Point(8, 8);
			this.lblListaSegmentos.Size      = new Size(500, 18);
			this.lblListaSegmentos.Font      = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold);

			// gridSegmentos
			this.gridSegmentos.Location      = new Point(8, 28);
			this.gridSegmentos.Size          = new Size(920, 220);
			this.gridSegmentos.Anchor        = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			this.gridSegmentos.ReadOnly      = true;
			this.gridSegmentos.CaptionVisible = false;
			this.gridSegmentos.BackgroundColor = SystemColors.Control;
			this.gridSegmentos.Click        += new EventHandler(this.gridSegmentos_Click);
			this.gridSegmentos.CurrentCellChanged += new EventHandler(this.gridSegmentos_CurrentCellChanged);

			// grpDetalle
			this.grpDetalle.Text     = "Configuración de cuentas contables";
			this.grpDetalle.Location = new Point(8, 258);
			this.grpDetalle.Size     = new Size(920, 270);
			this.grpDetalle.Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			this.grpDetalle.Enabled  = false;

			// lblSegmentoNombre
			this.lblSegmentoNombre.Text      = "Seleccione un segmento de la lista";
			this.lblSegmentoNombre.Location  = new Point(10, 22);
			this.lblSegmentoNombre.Size      = new Size(890, 20);
			this.lblSegmentoNombre.Font      = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold);
			this.lblSegmentoNombre.ForeColor = Color.DarkBlue;

			// --- Fila Cuenta Cartera (y1=52) ---
			this.lblCartera.Text     = "Cuenta Cartera:";
			this.lblCartera.Location = new Point(10, 55);
			this.lblCartera.Size     = new Size(110, 18);

			this.txtCarteraCodigo.Location  = new Point(125, 52);
			this.txtCarteraCodigo.Size      = new Size(100, 20);
			this.txtCarteraCodigo.CharacterCasing = CharacterCasing.Upper;
			this.txtCarteraCodigo.KeyDown  += new KeyEventHandler(this.txtCarteraCodigo_KeyDown);

			this.btnBuscarCartera.Text      = "...";
			this.btnBuscarCartera.Location  = new Point(228, 51);
			this.btnBuscarCartera.Size      = new Size(26, 22);
			this.btnBuscarCartera.Click    += new EventHandler(this.btnBuscarCartera_Click);

			this.txtCarteraNombre.Location  = new Point(258, 52);
			this.txtCarteraNombre.Size      = new Size(400, 20);
			this.txtCarteraNombre.ReadOnly  = true;
			this.txtCarteraNombre.BackColor = Color.WhiteSmoke;

			this.lblCarteraHint.Text        = "(F3 para buscar)";
			this.lblCarteraHint.Location    = new Point(665, 55);
			this.lblCarteraHint.Size        = new Size(120, 16);
			this.lblCarteraHint.ForeColor   = Color.Gray;
			this.lblCarteraHint.Font        = new Font("Microsoft Sans Serif", 7.5f, FontStyle.Italic);

			// --- Fila Cuenta Descuento (y2=88) ---
			this.lblDescuento.Text          = "Cuenta Descuento:";
			this.lblDescuento.Location      = new Point(10, 91);
			this.lblDescuento.Size          = new Size(110, 18);

			this.txtDescuentoCodigo.Location  = new Point(125, 88);
			this.txtDescuentoCodigo.Size      = new Size(100, 20);
			this.txtDescuentoCodigo.CharacterCasing = CharacterCasing.Upper;
			this.txtDescuentoCodigo.KeyDown  += new KeyEventHandler(this.txtDescuentoCodigo_KeyDown);

			this.btnBuscarDescuento.Text      = "...";
			this.btnBuscarDescuento.Location  = new Point(228, 87);
			this.btnBuscarDescuento.Size      = new Size(26, 22);
			this.btnBuscarDescuento.Click    += new EventHandler(this.btnBuscarDescuento_Click);

			this.txtDescuentoNombre.Location  = new Point(258, 88);
			this.txtDescuentoNombre.Size      = new Size(400, 20);
			this.txtDescuentoNombre.ReadOnly  = true;
			this.txtDescuentoNombre.BackColor = Color.WhiteSmoke;

			this.lblDescuentoHint.Text        = "(F3 para buscar)";
			this.lblDescuentoHint.Location    = new Point(665, 91);
			this.lblDescuentoHint.Size        = new Size(120, 16);
			this.lblDescuentoHint.ForeColor   = Color.Gray;
			this.lblDescuentoHint.Font        = new Font("Microsoft Sans Serif", 7.5f, FontStyle.Italic);

			// --- Fila Cuenta Recaudo (y3=124) ---
			this.lblRecaudo.Text            = "Cuenta Recaudo:";
			this.lblRecaudo.Location        = new Point(10, 127);
			this.lblRecaudo.Size            = new Size(110, 18);

			this.txtRecaudoCodigo.Location  = new Point(125, 124);
			this.txtRecaudoCodigo.Size      = new Size(100, 20);
			this.txtRecaudoCodigo.CharacterCasing = CharacterCasing.Upper;
			this.txtRecaudoCodigo.KeyDown  += new KeyEventHandler(this.txtRecaudoCodigo_KeyDown);

			this.btnBuscarRecaudo.Text      = "...";
			this.btnBuscarRecaudo.Location  = new Point(228, 123);
			this.btnBuscarRecaudo.Size      = new Size(26, 22);
			this.btnBuscarRecaudo.Click    += new EventHandler(this.btnBuscarRecaudo_Click);

			this.txtRecaudoNombre.Location  = new Point(258, 124);
			this.txtRecaudoNombre.Size      = new Size(400, 20);
			this.txtRecaudoNombre.ReadOnly  = true;
			this.txtRecaudoNombre.BackColor = Color.WhiteSmoke;

			this.lblRecaudoHint.Text        = "(F3 para buscar)";
			this.lblRecaudoHint.Location    = new Point(665, 127);
			this.lblRecaudoHint.Size        = new Size(120, 16);
			this.lblRecaudoHint.ForeColor   = Color.Gray;
			this.lblRecaudoHint.Font        = new Font("Microsoft Sans Serif", 7.5f, FontStyle.Italic);

			// --- Referencias de cuentas (yHints=158) ---
			this.lblRefs = new System.Windows.Forms.Label();
			this.lblRefs.Text      = "Referencia:  Cartera = cuenta del activo de cartera por segmento  |  " +
			                    "Descuento = cuenta de gasto por castigo/descuento  |  " +
			                    "Recaudo = cuenta de caja/banco donde ingresa el pago";
			this.lblRefs.Location  = new Point(10, 158);
			this.lblRefs.Size      = new Size(890, 32);
			this.lblRefs.ForeColor = Color.DimGray;
			this.lblRefs.Font      = new Font("Microsoft Sans Serif", 7.5f, FontStyle.Italic);

			// --- Activo (yAct=196) ---
			this.chkActivo.Text     = "Segmento activo (disponible en el proceso de cierre)";
			this.chkActivo.Location = new Point(10, 196);
			this.chkActivo.Size     = new Size(350, 20);
			this.chkActivo.Checked  = true;

			// --- Botones ---
			this.btnGuardar.Text      = "Guardar";
			this.btnGuardar.Location  = new Point(680, 194);
			this.btnGuardar.Size      = new Size(100, 26);
			this.btnGuardar.BackColor = Color.SteelBlue;
			this.btnGuardar.ForeColor = Color.White;
			this.btnGuardar.Click    += new EventHandler(this.btnGuardar_Click);

			this.btnCancelar.Text     = "Cancelar";
			this.btnCancelar.Location = new Point(788, 194);
			this.btnCancelar.Size     = new Size(100, 26);
			this.btnCancelar.Click   += new EventHandler(this.btnCancelar_Click);

			this.grpDetalle.Controls.AddRange(new Control[] {
				lblSegmentoNombre,
				lblCartera,   txtCarteraCodigo,   btnBuscarCartera,   txtCarteraNombre,   lblCarteraHint,
				lblDescuento, txtDescuentoCodigo, btnBuscarDescuento, txtDescuentoNombre, lblDescuentoHint,
				lblRecaudo,   txtRecaudoCodigo,   btnBuscarRecaudo,   txtRecaudoNombre,   lblRecaudoHint,
				lblRefs, chkActivo, btnGuardar, btnCancelar
			});

			// StatusBar
			this.statusBar.Panels.Add(this.sbpMensaje);
			this.statusBar.ShowPanels  = true;
			this.sbpMensaje.AutoSize   = StatusBarPanelAutoSize.Spring;
			this.sbpMensaje.Text       = "Listo";

			this.Controls.AddRange(new Control[] {
				lblListaSegmentos, gridSegmentos, grpDetalle, statusBar });

			// cdsSeteoF
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary        = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl     = "";
			this.cdsSeteoF.DataSetDef         = "dsSeteoF";
			this.cdsSeteoF.Locale             = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name               = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName    = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef          = null;
			this.cdsSeteoF.BeforeFill        += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);

			((System.ComponentModel.ISupportInitialize)(this.gridSegmentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbpMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);
		}
	}
}
