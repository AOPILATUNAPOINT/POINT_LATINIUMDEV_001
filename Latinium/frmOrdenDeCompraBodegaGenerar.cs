using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;

using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

using Latinium.Services.OrdenDeCompra.Bodega;

namespace Latinium
{
	/// <summary>
	/// DISTRIBUCIÓN POR BODEGAS
	/// </summary>
	public class frmOrdenDeCompraBodegaGenerar : System.Windows.Forms.Form
	{
		private System.ComponentModel.Container components = null;

		#region VARIABLES

		private int IdDetOrdenDeCompra = 0;

		public C1.Data.C1DataSet cdsSeteoF;

		private DataTable dtDetalle;

		private OrdenCompraBodegaGeneracionService service;

		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSourceBodega;

		#endregion

		#region CONTROLES


		private Label lblArticulo;

		private Label lblCodigo;

		private Label lblCantidadTotal;

		private Label lblDistribuido;

		private Label lblPendiente;

		private Label lblBuscar;

		private TextBox txtBuscar;

		private UltraGrid ultraGrid1;

		private Label lblTotalGenerar;

		private Button btnDistribuirTodo;

		private Button btnLimpiar;

		private Button btnGuardar;

		private Button btnCancelar;

		#endregion

		#region CONSTRUCTOR

		public frmOrdenDeCompraBodegaGenerar(
			int idDetOrdenDeCompra
			)
		{
			IdDetOrdenDeCompra =
				idDetOrdenDeCompra;

			InitializeComponent();
		}

		#endregion

		#region Dispose

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}

			base.Dispose(disposing);
		}

		#endregion

		#region DISEÑO

		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenDeCompraBodegaDetalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadAsignada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadGenerada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadTemporal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoProceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDeCompraBodegaDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadAsignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadGenerada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadTemporal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoProceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.ultraDataSourceBodega = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblArticulo = new System.Windows.Forms.Label();
			this.lblCodigo = new System.Windows.Forms.Label();
			this.lblCantidadTotal = new System.Windows.Forms.Label();
			this.lblDistribuido = new System.Windows.Forms.Label();
			this.lblPendiente = new System.Windows.Forms.Label();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.txtBuscar = new System.Windows.Forms.TextBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblTotalGenerar = new System.Windows.Forms.Label();
			this.btnDistribuirTodo = new System.Windows.Forms.Button();
			this.btnLimpiar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSourceBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSourceBodega
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSourceBodega.Band.Columns.AddRange(new object[] {
																																			ultraDataColumn1,
																																			ultraDataColumn2,
																																			ultraDataColumn3,
																																			ultraDataColumn4,
																																			ultraDataColumn5,
																																			ultraDataColumn6,
																																			ultraDataColumn7,
																																			ultraDataColumn8,
																																			ultraDataColumn9,
																																			ultraDataColumn10,
																																			ultraDataColumn11,
																																			ultraDataColumn12,
																																			ultraDataColumn13,
																																			ultraDataColumn14});
			// 
			// lblArticulo
			// 
			this.lblArticulo.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.lblArticulo.Location = new System.Drawing.Point(8, 8);
			this.lblArticulo.Name = "lblArticulo";
			this.lblArticulo.Size = new System.Drawing.Size(700, 18);
			this.lblArticulo.TabIndex = 1;
			this.lblArticulo.Text = "Artículo :";
			// 
			// lblCodigo
			// 
			this.lblCodigo.Location = new System.Drawing.Point(8, 32);
			this.lblCodigo.Name = "lblCodigo";
			this.lblCodigo.Size = new System.Drawing.Size(400, 18);
			this.lblCodigo.TabIndex = 2;
			this.lblCodigo.Text = "Código :";
			// 
			// lblCantidadTotal
			// 
			this.lblCantidadTotal.Location = new System.Drawing.Point(8, 64);
			this.lblCantidadTotal.Name = "lblCantidadTotal";
			this.lblCantidadTotal.Size = new System.Drawing.Size(250, 18);
			this.lblCantidadTotal.TabIndex = 3;
			this.lblCantidadTotal.Text = "Cantidad Total : 0";
			// 
			// lblDistribuido
			// 
			this.lblDistribuido.Location = new System.Drawing.Point(296, 64);
			this.lblDistribuido.Name = "lblDistribuido";
			this.lblDistribuido.Size = new System.Drawing.Size(250, 18);
			this.lblDistribuido.TabIndex = 4;
			this.lblDistribuido.Text = "Distribuido : 0";
			// 
			// lblPendiente
			// 
			this.lblPendiente.Location = new System.Drawing.Point(568, 64);
			this.lblPendiente.Name = "lblPendiente";
			this.lblPendiente.Size = new System.Drawing.Size(250, 18);
			this.lblPendiente.TabIndex = 5;
			this.lblPendiente.Text = "Pendiente : 0";
			// 
			// lblBuscar
			// 
			this.lblBuscar.Location = new System.Drawing.Point(8, 96);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(100, 20);
			this.lblBuscar.TabIndex = 6;
			this.lblBuscar.Text = "Buscar Bodega:";
			// 
			// txtBuscar
			// 
			this.txtBuscar.Location = new System.Drawing.Point(112, 96);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(250, 20);
			this.txtBuscar.TabIndex = 7;
			this.txtBuscar.Text = "";
			this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSourceBodega;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 82;
			ultraGridColumn5.Header.Caption = "Bodega";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 205;
			ultraGridColumn6.Header.Caption = "Artículo";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 157;
			ultraGridColumn7.Header.Caption = "Cantidad Asignada";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 204;
			ultraGridColumn8.Header.Caption = "Cantidad Generada";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 205;
			ultraGridColumn9.Header.Caption = "Saldo Pendiente";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 178;
			ultraGridColumn10.Header.Caption = "Cantidad Temporal";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 198;
			ultraGridColumn11.Header.Caption = "Estado Proceso";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 114;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 96;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
			this.ultraGrid1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
			appearance2.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellPadding = 3;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 8.5F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(992, 384);
			this.ultraGrid1.TabIndex = 8;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			// 
			// lblTotalGenerar
			// 
			this.lblTotalGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblTotalGenerar.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
			this.lblTotalGenerar.Location = new System.Drawing.Point(8, 536);
			this.lblTotalGenerar.Name = "lblTotalGenerar";
			this.lblTotalGenerar.Size = new System.Drawing.Size(400, 20);
			this.lblTotalGenerar.TabIndex = 9;
			this.lblTotalGenerar.Text = "Total a generar : 0";
			// 
			// btnDistribuirTodo
			// 
			this.btnDistribuirTodo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDistribuirTodo.Location = new System.Drawing.Point(8, 576);
			this.btnDistribuirTodo.Name = "btnDistribuirTodo";
			this.btnDistribuirTodo.Size = new System.Drawing.Size(120, 28);
			this.btnDistribuirTodo.TabIndex = 10;
			this.btnDistribuirTodo.Text = "Distribuir Todo";
			this.btnDistribuirTodo.Click += new System.EventHandler(this.btnDistribuirTodo_Click);
			// 
			// btnLimpiar
			// 
			this.btnLimpiar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnLimpiar.Location = new System.Drawing.Point(144, 576);
			this.btnLimpiar.Name = "btnLimpiar";
			this.btnLimpiar.Size = new System.Drawing.Size(100, 28);
			this.btnLimpiar.TabIndex = 11;
			this.btnLimpiar.Text = "Limpiar";
			this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Location = new System.Drawing.Point(760, 576);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(90, 28);
			this.btnGuardar.TabIndex = 12;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.Location = new System.Drawing.Point(870, 576);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(90, 28);
			this.btnCancelar.TabIndex = 13;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// frmOrdenDeCompraBodegaGenerar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 620);
			this.Controls.Add(this.lblArticulo);
			this.Controls.Add(this.lblCodigo);
			this.Controls.Add(this.lblCantidadTotal);
			this.Controls.Add(this.lblDistribuido);
			this.Controls.Add(this.lblPendiente);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblTotalGenerar);
			this.Controls.Add(this.btnDistribuirTodo);
			this.Controls.Add(this.btnLimpiar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.MinimumSize = new System.Drawing.Size(1000, 650);
			this.Name = "frmOrdenDeCompraBodegaGenerar";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Distribución por Bodegas";
			this.Load += new System.EventHandler(this.frmOrdenDeCompraBodegaGenerar_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSourceBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		#region LOAD

		private void frmOrdenDeCompraBodegaGenerar_Load(
			object sender,
			System.EventArgs e
			)
		{
			service =
				new OrdenCompraBodegaGeneracionService(
				cdsSeteoF
				);

			CargarDetalle();
		}

		#endregion

		#region CARGAR

		private void CargarDetalle()
		{
			try
			{
				dtDetalle =
					service.Consultar(
					IdDetOrdenDeCompra
					);

				// AÑADIR AQUÍ
				if (!dtDetalle.Columns.Contains("Sel"))
				{
					dtDetalle.Columns.Add("Sel", typeof(int));
				}

				foreach (DataRow r in dtDetalle.Rows)
				{
					if (r["Sel"] == DBNull.Value)
						r["Sel"] = 0;
				}

				this.ultraGrid1.DataSource =
					dtDetalle;

				ConfigurarGrid();

				CargarCabecera();

				ActualizarResumen();
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}
		#endregion

		#region CABECERA

		private void CargarCabecera()
		{
			if (dtDetalle.Rows.Count == 0)
				return;

			DataRow row =
				dtDetalle.Rows[0];

			this.lblArticulo.Text =
				"Artículo : " +
				row["Articulo"].ToString();

			this.lblCodigo.Text =
				"Código : " +
				row["Codigo"].ToString();
		}

		#endregion

		#region GRID

		private void ConfigurarGrid()
		{
			if (this.ultraGrid1.DisplayLayout.Bands.Count == 0)
				return;

			UltraGridBand band = this.ultraGrid1.DisplayLayout.Bands[0];

			// =========================
			// CONTROL DE EDICIÓN
			// =========================
			band.Columns["Sel"].CellActivation = Activation.AllowEdit;
			band.Columns["CantidadTemporal"].CellActivation = Activation.AllowEdit;

			band.Columns["NombreBodega"].CellActivation = Activation.NoEdit;
			band.Columns["CantidadAsignada"].CellActivation = Activation.NoEdit;
			band.Columns["CantidadGenerada"].CellActivation = Activation.NoEdit;
			band.Columns["SaldoPendiente"].CellActivation = Activation.NoEdit;

			// =========================
			// SOLO ENTEROS
			// =========================
			band.Columns["CantidadTemporal"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;

			// =========================
			// PERMISOS DEL GRID
			// =========================

			// NO PERMITIR ELIMINAR REGISTROS
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False;

			// NO PERMITIR AGREGAR NUEVAS FILAS DESDE EL GRID
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No;

			// PERMITIR ACTUALIZACIÓN SOLO DONDE SE PERMITA
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True;

			// =========================
			// COMPORTAMIENTO DE EDICIÓN
			// =========================
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = CellClickAction.Edit;

			// =========================
			// SELECCIÓN DE FILAS
			// =========================
			this.ultraGrid1.DisplayLayout.Override.RowSelectors = DefaultableBoolean.False;
		}

		#endregion

		#region GRID EVENTOS
		private bool _actualizandoGrid = false;
		private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e)
		{
			try
			{
				if (_actualizandoGrid)
					return;

				if (e.Cell == null || e.Cell.Row == null)
					return;

				if (e.Cell.Row.IsAddRow)
					return;

				if (e.Cell.Column.Key != "CantidadTemporal")
					return;

				int pendiente = 0;
				int temporal = 0;

				if (e.Cell.Row.Cells["SaldoPendiente"].Value != DBNull.Value)
					pendiente = Convert.ToInt32(e.Cell.Row.Cells["SaldoPendiente"].Value);

				if (e.Cell.Value != DBNull.Value)
				{
					try
					{
						temporal = Convert.ToInt32(e.Cell.Value);
					}
					catch
					{
						temporal = 0;
					}
				}

				int original = temporal;

				temporal = service.ValidarCantidad(temporal, pendiente);

				if (original != temporal)
				{
					MessageBox.Show(
						"La cantidad supera el saldo pendiente.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning
						);
				}

				_actualizandoGrid = true;   // BLOQUEO

				e.Cell.Value = temporal;

				if (e.Cell.Row.Cells.Exists("Sel"))
				{
					e.Cell.Row.Cells["Sel"].Value = (temporal > 0) ? 1 : 0;
				}

				PintarFila(e.Cell.Row);
				ActualizarResumen();

				_actualizandoGrid = false;  // LIBERACIÓN
			}
			catch (Exception ex)
			{
				_actualizandoGrid = false;
				MessageBox.Show(ex.Message);
			}
		}
		#endregion GRID EVENTOS

		#region PINTAR FILA

		private void PintarFila(
			UltraGridRow row
			)
		{
			int valor = 0;

			if (
				row.Cells["CantidadTemporal"].Value != DBNull.Value
				)
			{
				valor =
					Convert.ToInt32(
					row.Cells["CantidadTemporal"].Value
					);
			}

			if (valor > 0)
			{
				row.Appearance.BackColor =
					Color.Honeydew;
			}
			else
			{
				row.Appearance.ResetBackColor();
			}
		}

		#endregion

		#region BUSCAR

		private void txtBuscar_TextChanged(
			object sender,
			EventArgs e
			)
		{
			try
			{
				if (dtDetalle == null)
					return;

				string filtro =
					txtBuscar.Text.Trim().Replace("'", "''");

				if (filtro == "")
				{
					dtDetalle.DefaultView.RowFilter = "";
				}
				else
				{
					dtDetalle.DefaultView.RowFilter =
						string.Format(
						"NombreBodega LIKE '%{0}%'",
						filtro
						);
				}
			}
			catch
			{
			}
		}

		#endregion

		#region RESUMEN

		private void ActualizarResumen()
		{
			int totalGenerar = 0;
			int totalPendiente = 0;

			foreach (UltraGridRow row in ultraGrid1.Rows)
			{
				if (row == null || row.IsAddRow || row.Cells == null)
					continue;

				int temp = 0;
				int pend = 0;

				if (row.Cells.Exists("CantidadTemporal") &&
					row.Cells["CantidadTemporal"].Value != DBNull.Value)
					temp = Convert.ToInt32(row.Cells["CantidadTemporal"].Value);

				if (row.Cells["SaldoPendiente"].Value != DBNull.Value)
					pend = Convert.ToInt32(row.Cells["SaldoPendiente"].Value);

				totalGenerar += temp;
				totalPendiente += pend;
			}

			this.lblCantidadTotal.Text = "Cantidad Total : " + totalPendiente;
			this.lblDistribuido.Text = "Distribuido : " + totalGenerar;
			this.lblPendiente.Text = "Pendiente : " + (totalPendiente - totalGenerar);
			this.lblTotalGenerar.Text = "Total a generar : " + totalGenerar;
		}

		#endregion

		#region DISTRIBUIR TODO

		private void btnDistribuirTodo_Click(
			object sender,
			EventArgs e
			)
		{
			foreach (UltraGridRow row in ultraGrid1.Rows)
			{
				if (row == null || row.IsAddRow || row.Cells == null || row.Cells.Count == 0)
					continue;
				int pendiente = 0;

				if (
					row.Cells["SaldoPendiente"].Value != DBNull.Value
					)
				{
					pendiente =
						Convert.ToInt32(
						row.Cells["SaldoPendiente"].Value
						);
				}

				row.Cells["CantidadTemporal"].Value =
					pendiente;

				row.Cells["Sel"].Value = (pendiente > 0) ? 1 : 0;

				PintarFila(row);
			}

			ActualizarResumen();
		}

		#endregion

		#region LIMPIAR

		private void btnLimpiar_Click(
			object sender,
			EventArgs e
			)
		{
			foreach (UltraGridRow row in ultraGrid1.Rows)
			{
				if (row == null) continue;
				if (row.IsAddRow) continue;
				if (row.Cells == null) continue;
				row.Cells["CantidadTemporal"].Value = 0;

				row.Cells["Sel"].Value = 0;

				PintarFila(row);
			}

			ActualizarResumen();
		}

		#endregion

		#region GUARDAR

		private void btnGuardar_Click(
			object sender,
			System.EventArgs e
			)
		{
			try
			{
				this.ultraGrid1.UpdateData();

				service.GuardarTemporalMasivo(
					dtDetalle
					);

				MessageBox.Show(
					"Distribución guardada correctamente.",
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information
					);

				this.DialogResult =
					DialogResult.OK;

				this.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}

		#endregion

		#region CANCELAR

		private void btnCancelar_Click(
			object sender,
			System.EventArgs e
			)
		{
			this.DialogResult =
				DialogResult.Cancel;

			this.Close();
		}

		#endregion

		private void ultraGrid1_KeyPress(object sender, KeyPressEventArgs e)
		{
			UltraGridCell cell = this.ultraGrid1.ActiveCell;

			if (cell == null) return;

			if (cell.Column.Key == "CantidadTemporal")
			{
				// permitir solo números y control (backspace)
				if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
				{
					e.Handled = true;
				}
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.Key == "CantidadTemporal")
			{
				try
				{
					int valor = int.Parse(e.NewValue.ToString());
				}
				catch
				{
					e.Cancel = true;
				}
			}
		}

	}
}