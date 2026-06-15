using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmSeriales.
	/// </summary>
	public class frmSeriales : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private DevExpress.XtraEditors.RadioGroup rdgTipoConsulta;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnVer;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnExportar;

		private Acceso miAcceso;

		public frmSeriales()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Remate");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Remate");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDelSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Disponible");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDelSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Disponible");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeriales));
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.rdgTipoConsulta = new DevExpress.XtraEditors.RadioGroup();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 225;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 243;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Código";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 216;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Artículo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 300;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 227;
			ultraGridColumn7.Header.Caption = "Bodega Actual";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 226;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "R";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 24;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "SERIALES";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(10, 110);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1257, 265);
			this.ultraGrid1.TabIndex = 5;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// txtArticulo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance8;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.Location = new System.Drawing.Point(624, 46);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(269, 25);
			this.txtArticulo.TabIndex = 0;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance9.BackColor = System.Drawing.Color.White;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Appearance = appearance9;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 76;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 98;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 188;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Número";
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 107;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 107;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 263;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Width = 77;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Local";
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Width = 46;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Width = 239;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Sistema";
			ultraGridColumn19.Header.VisiblePosition = 10;
			ultraGridColumn19.Width = 185;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 55;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.Caption = "D";
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Width = 24;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand2.Header.Caption = "MOVIMIENTOS";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(10, 384);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1257, 219);
			this.ultraGrid2.TabIndex = 156;
			this.ultraGrid2.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_CellChange);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(bool);
			ultraDataColumn21.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 91);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1277, 10);
			this.groupBox1.TabIndex = 208;
			this.groupBox1.TabStop = false;
			// 
			// cmbBodega
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance16;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource3;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 60;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 224;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(278, 9);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(269, 25);
			this.cmbBodega.TabIndex = 3;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(211, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 18);
			this.label1.TabIndex = 210;
			this.label1.Text = "Bodega";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgTipoConsulta
			// 
			this.rdgTipoConsulta.EditValue = 0;
			this.rdgTipoConsulta.Location = new System.Drawing.Point(10, 9);
			this.rdgTipoConsulta.Name = "rdgTipoConsulta";
			// 
			// rdgTipoConsulta.Properties
			// 
			this.rdgTipoConsulta.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipoConsulta.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipoConsulta.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Serial - Artículo"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "N. Documento - Cliente")});
			this.rdgTipoConsulta.Size = new System.Drawing.Size(172, 64);
			this.rdgTipoConsulta.TabIndex = 2;
			this.rdgTipoConsulta.SelectedIndexChanged += new System.EventHandler(this.rdgTipoConsulta_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(211, 49);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 18);
			this.label2.TabIndex = 212;
			this.label2.Text = "Estado";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(566, 49);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 18);
			this.label3.TabIndex = 213;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(912, 45);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(90, 27);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbEstados
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstados.Appearance = appearance17;
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource4;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 224;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(278, 47);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(269, 25);
			this.cmbEstados.TabIndex = 4;
			this.cmbEstados.ValueMember = "idEstadoSerial";
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(624, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 24);
			this.lblContador.TabIndex = 661;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1018, 46);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(86, 26);
			this.btnExportar.TabIndex = 732;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmSeriales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(1276, 610);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.rdgTipoConsulta);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmSeriales";
			this.Text = "Movimiento de Seriales de Artículos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSeriales_KeyDown);
			this.Load += new System.EventHandler(this.frmSeriales_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}
		
		private void frmSeriales_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);

			miAcceso = new Acceso(cdsSeteoF, "0302");

			string sSQL = string.Format("Select idEstadoSerial, Estado From EstadosSeriales");
			this.cmbEstados.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
						
			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0) this.cmbBodega.Value = iBodega;				
			else 
			{
				this.cmbBodega.Value = 1;
				this.cmbBodega.Enabled = true;
			}
						
			if (miAcceso.Editar) this.ultraGrid2.DisplayLayout.Bands[0].Columns["Disponible"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
		}

		private void frmSeriales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	
		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null) return;
			if (this.ultraGrid1.ActiveRow.Cells["idSerial"].Value == DBNull.Value) return;

			string sSQL = string.Format("Exec ReporteDetalleSeriales {0}", (int)this.ultraGrid1.ActiveRow.Cells["idSerial"].Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ultraGrid2);			
		}
    bool bActualiza = false;
		private void ultraGrid2_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Row.Cells["idDetalleSerial"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Disponible")
				{
					bActualiza = true;
					if (!miAcceso.BEditarProcesado && miAcceso.Editar)
					{
						this.ultraGrid2.UpdateData();

						bool bDisponible = false;		
						if ((bool)e.Cell.Row.Cells["Disponible"].Value) bDisponible = true;
						else bDisponible = false;

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDisponible {0}, {1}", bDisponible, (int)e.Cell.Row.Cells["idDetalleSerial"].Value));
					}
					if (miAcceso.BEditarProcesado && miAcceso.Editar)
					{
						this.ultraGrid2.UpdateData();
						bool bDisponible = false;		
						if ((bool)e.Cell.Row.Cells["Disponible"].Value) bDisponible = true;
						else bDisponible = false;
						if (bDisponible)
						{
							int iValidaPayjoy = 0;
							iValidaPayjoy = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec ActualizarMovimientoPayJoy {0}, {1} ", (int)e.Cell.Row.Cells["idDetalleSerial"].Value, 1));
							if (iValidaPayjoy > 0)
							{
								#region Deiponible
								int iCont = 0;
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
								{
									if ((bool) dr.Cells["Disponible"].Value) iCont = iCont + 1;
								}
								if (iCont == 1 )
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDisponible {0}, {1}", bDisponible, (int)e.Cell.Row.Cells["idDetalleSerial"].Value));
								}
								else
								{
									e.Cell.Row.Cells["Disponible"].Value = (bool)e.Cell.OriginalValue;
									MessageBox.Show("Quite los vistos existentes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
								#endregion Deiponible
							}
							else
							{
								e.Cell.Row.Cells["Disponible"].Value = (bool)e.Cell.OriginalValue;
							}
						}
						else
						{
							#region nodisponible
							int iValidaPayjoyVenta = 0;
							iValidaPayjoyVenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec ActualizarMovimientoPayJoy {0}, {1} ", (int)e.Cell.Row.Cells["idDetalleSerial"].Value, 2));
							if (iValidaPayjoyVenta  == 1)
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDisponible {0}, {1}", bDisponible, (int)e.Cell.Row.Cells["idDetalleSerial"].Value));
								
							}
							else
							{
								e.Cell.Row.Cells["Disponible"].Value = (bool)e.Cell.OriginalValue;
								
							} 
							#endregion nodisponible
						}
					}
					bActualiza = false;
				}
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ReporteDetalleSeriales 0");
		
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;
			int iEstado = 0;
			if (this.cmbEstados.ActiveRow != null) iEstado = (int)this.cmbEstados.Value;

			string sSQL = string.Format("Exec ReporteSeriales '{0}', {1}, {2}, {3}", 
				this.txtArticulo.Text.ToString(), iEstado, iBodega, (int)this.rdgTipoConsulta.EditValue);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ultraGrid1);
			
			this.lblContador.Text = this.ultraGrid1.Rows.Count.ToString() + " Registros encontrados";
		}

		private void rdgTipoConsulta_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtArticulo.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstados.Focus();
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}
	}
}

