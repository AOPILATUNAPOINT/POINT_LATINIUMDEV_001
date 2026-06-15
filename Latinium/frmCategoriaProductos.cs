using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCategoriaProductos.
	/// </summary>
	public class frmCategoriaProductos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdCategoria;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCategoriaProductos;
		bool bActualiza = false;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCategoriaProductos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategoriaProductos");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCategoriaProductos));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCategoriaProductos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoProducto");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategoriaProductos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoProducto");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCategoriaProductos", -1, "cmbCategoriaProductos");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCategoriaProductos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.grdCategoria = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoriaProductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			this.SuspendLayout();
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1008, 32);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 24);
			this.btnExportar.TabIndex = 1050;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(376, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1051;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(80, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1052;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// cmbCategoriaProductos
			// 
			this.cmbCategoriaProductos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategoriaProductos.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCategoriaProductos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCategoriaProductos.DisplayMember = "TipoProducto";
			this.cmbCategoriaProductos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoriaProductos.Location = new System.Drawing.Point(568, 248);
			this.cmbCategoriaProductos.Name = "cmbCategoriaProductos";
			this.cmbCategoriaProductos.Size = new System.Drawing.Size(248, 96);
			this.cmbCategoriaProductos.TabIndex = 1053;
			this.cmbCategoriaProductos.ValueMember = "IdCategoriaProductos";
			this.cmbCategoriaProductos.Visible = false;
			// 
			// grdCategoria
			// 
			this.grdCategoria.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCategoria.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategoria.DisplayLayout.Appearance = appearance1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 83;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Código";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 215;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 550;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 188;
			ultraGridColumn7.Header.Caption = "Tipo Producto";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 131;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.grdCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategoria.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdCategoria.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 9F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCategoria.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.grdCategoria.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCategoria.Location = new System.Drawing.Point(0, 88);
			this.grdCategoria.Name = "grdCategoria";
			this.grdCategoria.Size = new System.Drawing.Size(1104, 296);
			this.grdCategoria.TabIndex = 1054;
			this.grdCategoria.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdCategoria_AfterCellUpdate_1);
			this.grdCategoria.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdCategoria_InitializeLayout);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 1056;
			this.label7.Text = "Buscar";
			this.label7.Click += new System.EventHandler(this.label7_Click);
			// 
			// txtBusca
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance8;
			this.txtBusca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusca.Location = new System.Drawing.Point(56, 8);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(300, 21);
			this.txtBusca.TabIndex = 1055;
			this.txtBusca.ValueChanged += new System.EventHandler(this.txtBusca_ValueChanged);
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.rdgProductoFacturable.EditValue = 1;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(496, 40);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "No Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Combos")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(472, 32);
			this.rdgProductoFacturable.TabIndex = 1059;
			this.rdgProductoFacturable.SelectedIndexChanged += new System.EventHandler(this.rdgProductoFacturable_SelectedIndexChanged);
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.rdgProductoExistencia.EditValue = 1;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(496, 8);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Con existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "En remates")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(472, 32);
			this.rdgProductoExistencia.TabIndex = 1058;
			this.rdgProductoExistencia.SelectedIndexChanged += new System.EventHandler(this.rdgProductoExistencia_SelectedIndexChanged);
			// 
			// frmCategoriaProductos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1120, 390);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtBusca);
			this.Controls.Add(this.cmbCategoriaProductos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.grdCategoria);
			this.Name = "frmCategoriaProductos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Categoria Productos";
			this.Load += new System.EventHandler(this.frmCategoriaProductos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoriaProductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ListarProductos '{0}',{1},{2}",this.txtBusca.Text.ToString(),(int)rdgProductoExistencia.EditValue, 
				(int)rdgProductoFacturable.EditValue);

			if (this.grdCategoria.Rows.Count > 0) this.grdCategoria.Focus(); else this.txtBusca.Focus();
			this.grdCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);


			
			this.lblContador.Text = this.grdCategoria.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void frmCategoriaProductos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200116");
			
			if (!Funcion.Permiso("1972", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Reporte de Frecuencia de bodegas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportarExcel) this.btnExportar.Enabled = true;


			this.btnVer_Click(sender, e);
			this.lblContador.Text = this.grdCategoria.Rows.Count + " REGISTROS ENCONTRADOS";

			this.cmbCategoriaProductos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT IdCategoriaProductos,TipoProducto FROM CategoriaProductos order by TipoProducto");	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.grdCategoria.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.grdCategoria);
		}
		private void ListaArticulosPrecios()
		{
		/*	this.grdCategoria.DataSource = null;
	
			string sSQL = string.Format("Exec CategoriaProductosBuffer '{0}',{1}",this.txtBusca.Text.ToString().Trim(),(int)optProductos.Value);

			if (this.grdCategoria.Rows.Count > 0) this.grdCategoria.Focus(); else this.txtBusca.Focus();
			this.grdCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.btnVer_Click(sender, e);
			
			this.lblContador.Text = this.grdCategoria.Rows.Count + " REGISTROS ENCONTRADOS";*/
			
		}
		private void grdCategoria_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void grdCategoria_AfterCellUpdate_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			/*if (bActualiza) return;

			if (e.Cell.Column.ToString() == "IdCategoriaProductos")
			{	
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Tipo del Producto?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE Articulo SET IdCategoriaProductos ={0} WHERE idArticulo ={1}",
					   (int)this.grdCategoria.ActiveRow.Cells["IdCategoriaProductos"].Value, (int)	this.grdCategoria.ActiveRow.Cells["idArticulo"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
				}
				else
				{						
					e.Cell.Row.Cells["IdCategoriaProductos"].Value = (int)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}	*/

			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "IdCategoriaProductos")
			{	
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Tipo de Producto?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE Articulo SET IdCategoriaProductos = {0} WHERE idArticulo = {1} ", (int)e.Cell.Row.Cells["IdCategoriaProductos"].Value ,(int)e.Cell.Row.Cells["idArticulo"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	
			
					this.btnVer_Click(sender, e);

				}
				else
				{						
					e.Cell.Row.Cells["IdCategoriaProductos"].Value = (int)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}	
		}

		private void label7_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtBusca_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void optBodegas_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender, e);
			//ListaArticulosPrecios();
		}

		private void rdgProductoFacturable_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		this.btnVer_Click(sender, e);
		}

		private void rdgProductoExistencia_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		this.btnVer_Click(sender, e);
		}

		private void grdCategoria_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
