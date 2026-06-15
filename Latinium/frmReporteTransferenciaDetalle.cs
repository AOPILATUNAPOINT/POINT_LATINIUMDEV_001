using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Latinium.Services.Transferencia;

namespace Latinium
{
	public class frmReporteTransferenciaDetalle : System.Windows.Forms.Form
	{
		private System.ComponentModel.Container components = null;

		private DateTimePicker dtpDesde;
		private DateTimePicker dtpHasta;
		private ComboBox cmbGrupo;
		private ComboBox cmbArticulo;
		private Button btnBuscar;
		private DataGrid dgReporte;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ComboBox cmbBodega;
		private Button btnExportar;

		private string _connectionString = "";

		public frmReporteTransferenciaDetalle()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.dtpDesde = new System.Windows.Forms.DateTimePicker();
			this.dtpHasta = new System.Windows.Forms.DateTimePicker();
			this.cmbGrupo = new System.Windows.Forms.ComboBox();
			this.cmbArticulo = new System.Windows.Forms.ComboBox();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.dgReporte = new System.Windows.Forms.DataGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBodega = new System.Windows.Forms.ComboBox();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dgReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpDesde
			// 
			this.dtpDesde.Location = new System.Drawing.Point(20, 20);
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.Size = new System.Drawing.Size(120, 20);
			this.dtpDesde.TabIndex = 0;
			// 
			// dtpHasta
			// 
			this.dtpHasta.Location = new System.Drawing.Point(160, 20);
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.Size = new System.Drawing.Size(120, 20);
			this.dtpHasta.TabIndex = 1;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.Location = new System.Drawing.Point(470, 20);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(150, 21);
			this.cmbGrupo.TabIndex = 3;
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cmbArticulo.Location = new System.Drawing.Point(640, 20);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(264, 21);
			this.cmbArticulo.TabIndex = 4;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnBuscar.Location = new System.Drawing.Point(1032, 20);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 5;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// dgReporte
			// 
			this.dgReporte.AlternatingBackColor = System.Drawing.Color.LightGray;
			this.dgReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgReporte.CaptionText = "Resultado del Reporte";
			this.dgReporte.DataMember = "";
			this.dgReporte.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgReporte.Location = new System.Drawing.Point(20, 60);
			this.dgReporte.Name = "dgReporte";
			this.dgReporte.ReadOnly = true;
			this.dgReporte.Size = new System.Drawing.Size(1086, 480);
			this.dgReporte.TabIndex = 6;
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
			// cmbBodega
			// 
			this.cmbBodega.Location = new System.Drawing.Point(300, 20);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(150, 21);
			this.cmbBodega.TabIndex = 7;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Location = new System.Drawing.Point(936, 20);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 8;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmReporteTransferenciaDetalle
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1128, 566);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtpDesde);
			this.Controls.Add(this.dtpHasta);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.dgReporte);
			this.MinimumSize = new System.Drawing.Size(900, 500);
			this.Name = "frmReporteTransferenciaDetalle";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte de Transferencias";
			this.Resize += new System.EventHandler(this.frmReporteTransferenciaDetalle_Resize);
			this.Load += new System.EventHandler(this.frmReporteTransferenciaDetalle_Load);
			((System.ComponentModel.ISupportInitialize)(this.dgReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}

		private void UnloadMe()
		{
			this.Close();
		}

		// ============================
		// LOAD
		// ============================
		private void frmReporteTransferenciaDetalle_Load(object sender, EventArgs e)
		{

			if (!Funcion.Permiso("2213", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a línea de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			try
			{
				_connectionString = Funcion.ObtenerCadenaConexionEmpresa(
					MenuLatinium.stIdDB, 
					frmRequerimiento.sconexionInicio, 
					frmSeguimiento.sconexionEmpresaSeguimiento
					);

				ReporteTransferenciaService service = new ReporteTransferenciaService(_connectionString);

				// Fechas
				dtpDesde.Value = DateTime.Now.AddDays(-7);
				dtpHasta.Value = DateTime.Now;

				// ============================
				// BODEGAS DESDE SERVICE
				// ============================
				int bodegaAsignada = 0;

				DataTable dtBodega = service.ObtenerBodegasPorUsuario(
					MenuLatinium.IdUsuario,
					1,
					DateTime.Today,
					false,
					out bodegaAsignada
					);

				cmbBodega.DataSource = dtBodega;
				cmbBodega.DisplayMember = "Nombre";
				cmbBodega.ValueMember = "Bodega";

				// LÓGICA DE NEGOCIO
				if (bodegaAsignada > 0)
				{
					cmbBodega.SelectedValue = bodegaAsignada;
					cmbBodega.Enabled = false;
				}
				else
				{
					cmbBodega.Enabled = true;
				}

				// ============================
				// GRUPOS
				// ============================
				cmbGrupo.DataSource = service.ObtenerGrupos();
				cmbGrupo.DisplayMember = "Grupo";
				cmbGrupo.ValueMember = "idGrupoArticulo";

				// ============================
				// ARTICULOS
				// ============================
				cmbArticulo.DataSource = service.ObtenerArticulos();
				cmbArticulo.DisplayMember = "Articulo";
				cmbArticulo.ValueMember = "idArticulo";

				AjustarControles();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al cargar pantalla: " + ex.Message);
			}
		}

		// ============================
		// BUSCAR
		// ============================
		private void btnBuscar_Click(object sender, EventArgs e)
		{
			try
			{
				if (dtpDesde.Value > dtpHasta.Value)
				{
					MessageBox.Show("La fecha desde no puede ser mayor que la hasta");
					return;
				}

				ReporteTransferenciaService service = new ReporteTransferenciaService(_connectionString);

				DataTable dt = service.ObtenerReporte(
					dtpDesde.Value,
					dtpHasta.Value,
					-1,
					Convert.ToInt32(cmbBodega.SelectedValue),
					Convert.ToInt32(cmbGrupo.SelectedValue),
					Convert.ToInt32(cmbArticulo.SelectedValue)
					);

				dgReporte.DataSource = dt;

				ConfigurarGrid(dt);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al consultar: " + ex.Message);
			}
		}

		// ============================
		// RESPONSIVE
		// ============================
		private void frmReporteTransferenciaDetalle_Resize(object sender, EventArgs e)
		{
			AjustarControles();
		}

		private void AjustarControles()
		{
			int margen = 20;
			int espacio = 10;

			// ============================
			// BOTONES (DERECHA)
			// ============================

			btnBuscar.Left = this.ClientSize.Width - btnBuscar.Width - margen;
			btnExportar.Left = btnBuscar.Left - btnExportar.Width - espacio;

			// ============================
			// ARTICULO dinámico
			// ============================

			cmbArticulo.Width = btnExportar.Left - cmbArticulo.Left - espacio;

			// ============================
			// CONTROLES IZQUIERDA (ORDENADOS)
			// ============================

			dtpDesde.Left = margen;
			dtpHasta.Left = dtpDesde.Right + espacio;
			cmbBodega.Left = dtpHasta.Right + espacio;
			cmbGrupo.Left = cmbBodega.Right + espacio;

			// ============================
			// VALIDACIÓN RESPONSIVE (EVITA QUE SE ROMPA)
			// ============================

			if (cmbGrupo.Right + 200 > cmbArticulo.Left)
			{
				cmbGrupo.Width = Math.Max(100, cmbArticulo.Left - cmbGrupo.Left - espacio);
			}

			// ============================
			// GRID
			// ============================

			dgReporte.Width = this.ClientSize.Width - (margen * 2);
			dgReporte.Height = this.ClientSize.Height - dgReporte.Top - margen;
		}
		// ============================
		// GRID
		// ============================
		private void ConfigurarGrid(DataTable dt)
		{
			try
			{
				dgReporte.TableStyles.Clear();

				DataGridTableStyle tableStyle = new DataGridTableStyle();
				tableStyle.MappingName = dt.TableName;

				AgregarColumna(tableStyle, "Fecha", "Fecha", 120);
				AgregarColumna(tableStyle, "Serial", "Serial Entregado", 150);
				AgregarColumna(tableStyle, "Peso", "Peso", 80);
				AgregarColumna(tableStyle, "CodigoProducto", "Código", 100);
				AgregarColumna(tableStyle, "Modelo", "Modelo", 120);
				AgregarColumna(tableStyle, "Marca", "Marca", 100);
				AgregarColumna(tableStyle, "Grupo", "Grupo", 100);
				AgregarColumna(tableStyle, "SubGrupo", "SubGrupo", 100);
				AgregarColumna(tableStyle, "BodegaOrigen", "Bodega Origen", 120);
				AgregarColumna(tableStyle, "BodegaDestino", "Bodega Destino", 120);
				AgregarColumna(tableStyle, "Solicita", "Solicita", 120);
				AgregarColumna(tableStyle, "CodigoTransferencia", "Transferencia", 100);
				AgregarColumna(tableStyle, "Articulo", "Artículo", 250);
				AgregarColumna(tableStyle, "Cantidad", "Cantidad", 80);
				AgregarColumna(tableStyle, "CantidadEntregada", "Cantidad Entregada", 100);
				AgregarColumna(tableStyle, "Nota", "Nota", 200);
				AgregarColumna(tableStyle, "UsuarioModifica", "Usuario Modifica", 130);
				AgregarColumna(tableStyle, "Costo", "Costo", 80);
				AgregarColumna(tableStyle, "Estado", "Estado", 100);

				dgReporte.TableStyles.Add(tableStyle);
			}
			catch { }
		}

		private void AgregarColumna(DataGridTableStyle tableStyle, string mapping, string header, int width)
		{
			DataGridTextBoxColumn col = new DataGridTextBoxColumn();
			col.MappingName = mapping;
			col.HeaderText = header;
			col.Width = width;
			tableStyle.GridColumnStyles.Add(col);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExportar_Click(object sender, EventArgs e)
		{
			try
			{
				if (dgReporte.DataSource == null)
				{
					MessageBox.Show("No hay datos para exportar");
					return;
				}

				DataTable dt = dgReporte.DataSource as DataTable;

				if (dt == null || dt.Rows.Count == 0)
				{
					MessageBox.Show("No hay datos para exportar");
					return;
				}

				// ============================
				// SAVE FILE DIALOG AQUÍ (CORRECTO)
				// ============================

				SaveFileDialog dlg = new SaveFileDialog();
				dlg.Filter = "Archivo Excel (*.xls)|*.xls";
				dlg.Title = "Guardar reporte";
				dlg.FileName = "ReporteTransferencias_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

				if (dlg.ShowDialog() != DialogResult.OK)
					return;

				string ruta = dlg.FileName;

				// ============================
				// LLAMADA AL SERVICE
				// ============================

				ReporteTransferenciaService service = new ReporteTransferenciaService(_connectionString);
				service.ExportarExcelHtml(dt, ruta);

				MessageBox.Show("Archivo exportado en: " + ruta);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al exportar: " + ex.Message);
			}
		}
	}
}