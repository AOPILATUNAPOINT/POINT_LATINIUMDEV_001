using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPreciadorRemates.
	/// </summary>
	public class frmPreciadorRemates : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblVendedor;
		private DevExpress.XtraEditors.RadioGroup rdgPreciador;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPreciadorRemates()
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

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPreciadorRemates));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelPrecios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Llegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Normal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Remate");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuota");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SelPrecios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Llegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio_Normal");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio_Remate");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuota");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.btnExcel = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.rdgPreciador = new DevExpress.XtraEditors.RadioGroup();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.rdgPreciador.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1136, 39);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 295;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(144, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 17);
			this.label1.TabIndex = 291;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(144, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 290;
			this.lblVendedor.Text = "Subgrupo";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgPreciador
			// 
			this.rdgPreciador.EditValue = 1;
			this.rdgPreciador.Location = new System.Drawing.Point(8, 8);
			this.rdgPreciador.Name = "rdgPreciador";
			// 
			// rdgPreciador.Properties
			// 
			this.rdgPreciador.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgPreciador.Properties.Appearance.Options.UseBackColor = true;
			this.rdgPreciador.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Grande"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Mediano"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Pequeño")});
			this.rdgPreciador.Size = new System.Drawing.Size(128, 56);
			this.rdgPreciador.TabIndex = 294;
			this.rdgPreciador.SelectedIndexChanged += new System.EventHandler(this.rdgPreciador_SelectedIndexChanged);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(424, 40);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 293;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1048, 40);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 292;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtBuscar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(216, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 288;
			// 
			// cmbSubGrupo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance2;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 368;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubGrupo.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False;
			this.cmbSubGrupo.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(216, 8);
			this.cmbSubGrupo.MaxDropDownItems = 10;
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(288, 21);
			this.cmbSubGrupo.TabIndex = 287;
			this.cmbSubGrupo.ValueMember = "idSubGrupo";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn13});
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Serie";
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 154;
			ultraGridColumn4.Header.Caption = "...";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 20;
			ultraGridColumn5.Header.Caption = "I";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 20;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 182;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 277;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 148;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Ingreso";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 79;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "% Desc";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 69;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance6;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 79;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance7;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Precio Normal";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 88;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance8;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Precio Remate";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 88;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance9;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Valor Cuota";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 80;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1206, 272);
			this.ultraGrid1.TabIndex = 296;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1224, 8);
			this.groupBox1.TabIndex = 297;
			this.groupBox1.TabStop = false;
			// 
			// optFiltro
			// 
			this.optFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance16;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance17;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Pendiente Impresión";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Impresos";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(675, 6);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(224, 24);
			this.optFiltro.TabIndex = 298;
			this.optFiltro.Text = "Pendiente Impresión";
			this.optFiltro.Visible = false;
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(520, 41);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 299;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(931, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 17);
			this.label2.TabIndex = 701;
			this.label2.Text = "Bodega";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance18;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 180;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(995, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(216, 22);
			this.cmbBodega.TabIndex = 700;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// frmPreciadorRemates
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1220, 366);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.rdgPreciador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmPreciadorRemates";
			this.Text = "Impresión de Precios de Remate";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPreciadorRemates_KeyDown);
			this.Load += new System.EventHandler(this.frmPreciadorRemates_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgPreciador.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void frmPreciadorRemates_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0731");

			if (miAcceso.Buscar) this.optFiltro.Visible = true;
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			if (miAcceso.CambiarDescuento)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescPorc"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Precio_Remate"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubGrupo, SubGrupo From ArticuloSubGrupo Where idPreciador = 1 Order By SubGrupo Asc");			

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Codigo, Nombre, Bodega From Bodega Where Bodega In (36, 81, 94)");

			string sSQL = string.Format("Exec RetornaSegBodegaFactura {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 1) this.cmbBodega.Value = iBodega; 
			else
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 36;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta()
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;

			int idSubgrupo = 0;
			if (this.cmbSubGrupo.ActiveRow != null) idSubgrupo = (int)this.cmbSubGrupo.Value;

			string sSQL = string.Format("Exec RMListaImpresionDePrecios {0}, {1}, '{2}', {3}, {4}", 
				idSubgrupo, (int)this.rdgPreciador.EditValue, this.txtBuscar.Text.ToString(), (bool)this.optFiltro.Value, (int)this.cmbBodega.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SelPrecios")
			{
				if (e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value != System.DBNull.Value)
				{
					this.ultraGrid1.UpdateData();
					
					string sSQL = string.Format("Exec SeleccionaArticulosImpPreciosRemate {0}, {1}, {2}", 
						(int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, (bool)e.Cell.Row.Cells["SelPrecios"].Value, (bool)this.optFiltro.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}
			}

			if (e.Cell.Column.ToString() == "Impreso")
			{
				if (e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value != System.DBNull.Value)
				{
					this.ultraGrid1.UpdateData();
					
					string sSQL = string.Format("Exec SeleccionaArticulosImpPreciosRemate {0}, {1}, {2}", 
						(int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value, (bool)e.Cell.Row.Cells["Impreso"].Value, (bool)this.optFiltro.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros para Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @Res Int = 0 If Not Exists(Select * From DTDetalleIngresoRetiros d Inner Join Articulo a On a.idArticulo = d.idArticuloMod Where Impreso = 0 And SelPrecios = 1 And idPreciador = {0}) Set @Res = 1 Select @Res", (int)this.rdgPreciador.EditValue)) == 1)
			{
				MessageBox.Show("Seleccione los Articulos para Imprimir los Precios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						
			string sDir = "";
			string sDirPCG = "";
			string sDirPCM = "";
			string sDirPCP = "";			

			string sSQL = string.Format("Select DirectorioPrecios, DirectorioPCM, DirectorioPCP From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				sDirPCG = dr.GetString(0);
				sDirPCM = dr.GetString(1);
				sDirPCP = dr.GetString(2);				
			}
			dr.Close();

			if ((int)this.rdgPreciador.EditValue == 1) sDir = sDirPCG;
			if ((int)this.rdgPreciador.EditValue == 2) sDir = sDirPCM;
			if ((int)this.rdgPreciador.EditValue == 3) sDir = sDirPCP;
				
			Process.Start(sDir, "");

			MessageBox.Show("Imprimiendo precios\n\npor favor espere", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DTDetalleIngresoRetiros Set Impreso = 1, SelPrecios = 0 Where SelPrecios = 1 And Impreso = 0 And idDetalleIngresoRetiro = {0}", 
					(int)dr1.Cells["idDetalleIngresoRetiro"].Value));
			}
			this.btnBuscar_Click(sender, e);
		}

		private void frmPreciadorRemates_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void rdgPreciador_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["SelPrecios"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			this.btnImprimir.Enabled = true;

			if ((int)this.rdgPreciador.EditValue == 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SelPrecios"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.btnImprimir.Enabled = false;
			}
			
			Consulta();
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.btnImprimir.Enabled = false;

			if (!(bool)this.optFiltro.Value) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SelPrecios"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].Hidden = true;

				if (miAcceso.Imprimir) if ((int)this.rdgPreciador.EditValue > 0) this.btnImprimir.Enabled = true;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SelPrecios"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].Hidden = false;
				
				if (miAcceso.ReImprimir) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;				
			}

			Consulta();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "DescPorc", 2);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Precio_Remate")
			{	
				decimal dNuevoPrecioSinIVA = Convert.ToDecimal(e.Cell.Row.Cells["Precio_Remate"].Value)/1.12m;

				e.Cell.Row.Cells["DescPorc"].Value = - ((dNuevoPrecioSinIVA * 100.00m)/Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - 100.00m);
			}

			if (e.Cell.Column.ToString() == "DescPorc")
			{				
				decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DTDetalleIngresoRetiros Set Descuento = {0} Where idDetalleIngresoRetiro = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["DescPorc"].Value), (int)e.Cell.Row.Cells["idDetalleIngresoRetiro"].Value));

				e.Cell.Row.Cells["Descuento"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescPorc"].Value)/100.00m;
				e.Cell.Row.Cells["Precio_Normal"].Value = Decimal.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * dIVA, 2);
				e.Cell.Row.Cells["Precio_Remate"].Value = 
					Decimal.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - ((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescPorc"].Value))/100.00m)) * dIVA), 2);				
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}
	}
}
