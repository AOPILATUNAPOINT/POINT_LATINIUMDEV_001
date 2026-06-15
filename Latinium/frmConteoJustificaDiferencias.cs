using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmConteoJustificaDiferencias.
	/// </summary>
	public class frmConteoJustificaDiferencias : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInforme;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.ComboBox cboFecha;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConteoJustificaDiferencias()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCierreConteoInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConteoExistencia");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoJust", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadJustificado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrar");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasCierre");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cobrar", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cobrar", 9, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCierreConteoInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConteoExistencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoJust");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadJustificado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasCierre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoJust");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConteoJustificaDiferencias));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.uGridInforme = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.label7 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.label55 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cboFecha = new System.Windows.Forms.ComboBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInforme
			// 
			this.uGridInforme.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInforme.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInforme.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInforme.DisplayLayout.Appearance = appearance1;
			this.uGridInforme.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 152;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 111;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 35;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 44;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Articulo";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Header.Caption = "Diferencias";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 120;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Header.Caption = "Cant Just";
			ultraGridColumn9.Header.VisiblePosition = 8;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 250;
			ultraGridColumn12.Header.Caption = "Notas del Cierre";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 266;
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
																										 ultraGridColumn12});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.uGridInforme.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInforme.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInforme.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridInforme.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInforme.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridInforme.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridInforme.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInforme.Location = new System.Drawing.Point(8, 96);
			this.uGridInforme.Name = "uGridInforme";
			this.uGridInforme.Size = new System.Drawing.Size(1136, 208);
			this.uGridInforme.TabIndex = 368;
			this.uGridInforme.DoubleClick += new System.EventHandler(this.uGridInforme_DoubleClick);
			this.uGridInforme.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInforme_KeyDown);
			this.uGridInforme.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInforme_AfterCellUpdate);
			this.uGridInforme.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInforme_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn12});
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 139;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 278;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(640, 160);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(280, 112);
			this.cmbEstados.TabIndex = 375;
			this.cmbEstados.ValueMember = "idEstadoJust";
			this.cmbEstados.Visible = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(42, 16);
			this.label7.TabIndex = 378;
			this.label7.Text = "Artículo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArticulo
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance10;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(72, 40);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(264, 22);
			this.txtArticulo.TabIndex = 376;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(360, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 377;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(936, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 379;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Location = new System.Drawing.Point(8, 8);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(41, 16);
			this.label55.TabIndex = 383;
			this.label55.Text = "Bodega";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(360, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 382;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 180;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 21);
			this.cmbBodega.TabIndex = 380;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(1024, 40);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(120, 23);
			this.btnGuardar.TabIndex = 384;
			this.btnGuardar.Text = "Cerrar Inventario";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// cboFecha
			// 
			this.cboFecha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboFecha.Location = new System.Drawing.Point(416, 8);
			this.cboFecha.MaxDropDownItems = 15;
			this.cboFecha.Name = "cboFecha";
			this.cboFecha.Size = new System.Drawing.Size(112, 21);
			this.cboFecha.TabIndex = 385;
			this.cboFecha.SelectedIndexChanged += new System.EventHandler(this.cboFecha_SelectedIndexChanged);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(768, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 22);
			this.lblContador.TabIndex = 386;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1152, 8);
			this.groupBox1.TabIndex = 387;
			this.groupBox1.TabStop = false;
			// 
			// frmConteoJustificaDiferencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1152, 310);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cboFecha);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.uGridInforme);
			this.Name = "frmConteoJustificaDiferencias";
			this.Text = "Cierre Conteo De Inventario";
			this.Load += new System.EventHandler(this.frmConteoJustificaDiferencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridInforme.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridInforme, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;

			if (this.cboFecha.Items.Count == 0)	
			{
				MessageBox.Show("No Hay Inventarios Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.cboFecha.Text.Length == 0)	
			{
				MessageBox.Show("Seleccione Una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			DateTime dtFecha = DateTime.Parse(this.cboFecha.Text);

			string sSQL = string.Format("Exec ReporteJustificaDiferenciasInventario '{0}', '{1}', {2}", this.txtArticulo.Text.ToString(), dtFecha.ToString("yyyyMMdd"), iBodega);
			this.uGridInforme.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	
		
			this.lblContador.Text = this.uGridInforme.Rows.Count + " REGISTROS";
		}

		private void frmConteoJustificaDiferencias_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			int iBodega = 0;

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0)
			{
				this.cmbBodega.Value = iBodega;
				this.cmbBodega.Enabled = false;
			}
			else this.cmbBodega.Value = 1;

			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEstado As idEstadoJust, Estado From EstadoAjusteConteo Where idEstado In (3, 5, 6, 7) Order By idEstado");		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida Estados
			int iContJustificado = 0;
			int iContPerdida = 0;
			int iContCobro = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
			{
				if ((int)dr.Cells["idEstadoJust"].Value == 5) iContJustificado++;
				if ((int)dr.Cells["idEstadoJust"].Value == 6) iContCobro++;
				if ((int)dr.Cells["idEstadoJust"].Value == 7) iContPerdida++;
			}
			#endregion Valida Estados

			if (iContJustificado == 0 && iContCobro == 0 && iContPerdida == 0)
			{
				MessageBox.Show("Cambie los Estados Correspondientes para Crear los Egresos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				DateTime dtFecha = DateTime.Parse(this.cboFecha.Text);
				string sMensaje = "";
				
				#region Validacion
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
				{
					if ((int)dr.Cells["idEstadoJust"].Value == 5 || (int)dr.Cells["idEstadoJust"].Value == 6 || (int)dr.Cells["idEstadoJust"].Value == 7)
					{
						if ((int)dr.Cells["CantidadJustificado"].Value == 0)
						{							
							if ((int)dr.Cells["idEstadoJust"].Value == 5) sMensaje = " A JUSTIFICAR ";
							else if ((int)dr.Cells["idEstadoJust"].Value == 6) sMensaje = " A COBRAR ";
							else if ((int)dr.Cells["idEstadoJust"].Value == 7) sMensaje = " DE LA PERDIDA ";

							MessageBox.Show(string.Format("La Cantidad '{0}' Debe ser Mayor a CERO", sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							

							this.uGridInforme.ActiveRow = dr;
							this.uGridInforme.ActiveRow.Selected = true;

							return;
						}
					}

					if ((int)dr.Cells["idEstadoJust"].Value == 6)
					{
						if ((decimal)dr.Cells["Cobrar"].Value == 0.00m)
						{
							MessageBox.Show("Ingrese el Valor del Cobro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

							this.uGridInforme.ActiveRow = dr;
							this.uGridInforme.ActiveRow.Selected = true;

							return;
						}
					}
				}
				#endregion Validacion

				int idEgresoJustificado = 0;
				int idEgresoCobrado = 0;
				int idEgresoPerdida = 0;

				#region Egreso Justificados
				if (iContJustificado > 0)
				{	
					string sSQLJustificado = string.Format("Exec EgresoCierreInventario '{0}', {1}, 5", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);	
					idEgresoJustificado = Funcion.iEscalarSQL(cdsSeteoF, sSQLJustificado, true);

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
					{
						if ((int)dr.Cells["idEstadoJust"].Value == 5)
						{
							string sSQLDetalleJustificado = string.Format("Exec DetalleEgresoCierreInventario {0}, {1}, 5", (int)dr.Cells["idLogAjusteCantidades"].Value, idEgresoJustificado);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleJustificado);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set Cerrados = Cerrados + {2}, Justificado = Justificado + {2}, idEgreso = {0} Where idLogAjusteCantidades = {1}", idEgresoCobrado, (int)dr.Cells["idLogAjusteCantidades"].Value, (int)dr.Cells["CantidadJustificado"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set idEstadoJust = Case When (ABS(ConteoExistencia) - CantidadJustifica) <> Cerrados Then 3 Else idEstadoJust End, Cerrado = Case When (ABS(ConteoExistencia) - CantidadJustifica) = Cerrados Then 1 Else 0 End Where idLogAjusteCantidades = {0}", (int)dr.Cells["idLogAjusteCantidades"].Value));
						}
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CreaAsiEgreso {0}", idEgresoJustificado));
				}
				#endregion Egreso Justificados	

				#region Egreso Cobro
				if (iContCobro > 0)
				{
					string sSQLCobro = string.Format("Exec EgresoCierreInventario '{0}', {1}, 6", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);	
					idEgresoCobrado = Funcion.iEscalarSQL(cdsSeteoF, sSQLCobro, true);
				
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
					{
						if ((int)dr.Cells["idEstadoJust"].Value == 6) 
						{
							string sSQLDetalleCobro = string.Format("Exec DetalleEgresoCierreInventario {0}, {1}, 6", (int)dr.Cells["idLogAjusteCantidades"].Value, idEgresoCobrado);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleCobro);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set Cerrados = Cerrados + {2}, Cobrado = Cobrado + {2}, idEgreso = {0} Where idLogAjusteCantidades = {1}", idEgresoCobrado, (int)dr.Cells["idLogAjusteCantidades"].Value, (int)dr.Cells["CantidadJustificado"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set idEstadoJust = Case When (ABS(ConteoExistencia) - CantidadJustifica) <> Cerrados Then 3 Else idEstadoJust End, Cerrado = Case When (ABS(ConteoExistencia) - CantidadJustifica) = Cerrados Then 1 Else 0 End Where idLogAjusteCantidades = {0}", (int)dr.Cells["idLogAjusteCantidades"].Value));
						}
					}				

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AsientoEgresoCierreInventario {0}", idEgresoCobrado));
				}
				#endregion Egreso Cobro

				#region Egreso Perdida
				if (iContPerdida > 0)
				{
					string sSQLPerdida = string.Format("Exec EgresoCierreInventario '{0}', {1}, 7", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value);	
					idEgresoPerdida = Funcion.iEscalarSQL(cdsSeteoF, sSQLPerdida, true);

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
					{						
						if ((int)dr.Cells["idEstadoJust"].Value == 7) 
						{
							string sSQLDetallePerdida = string.Format("Exec DetalleEgresoCierreInventario {0}, {1}, 7", (int)dr.Cells["idLogAjusteCantidades"].Value, idEgresoPerdida);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLDetallePerdida);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set Cerrados = Cerrados + {2}, Perdida = Perdida + {2}, idEgreso = {0} Where idLogAjusteCantidades = {1}", idEgresoCobrado, (int)dr.Cells["idLogAjusteCantidades"].Value, (int)dr.Cells["CantidadJustificado"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set idEstadoJust = Case When (ABS(ConteoExistencia) - CantidadJustifica) <> Cerrados Then 3 Else idEstadoJust End, Cerrado = Case When (ABS(ConteoExistencia) - CantidadJustifica) = Cerrados Then 1 Else 0 End Where idLogAjusteCantidades = {0}", (int)dr.Cells["idLogAjusteCantidades"].Value));
						}
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AsientoEgresoCierreInventario {0}", idEgresoPerdida));
				}
				#endregion Egreso Perdida
				
				MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select Fecha From CierreConteoInventario Where Bodega = {0} And Estado = 2 Order By Fecha", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
			this.cboFecha.Items.Clear();

			DateTime Fecha = DateTime.Today;
			
			while (dr.Read())
			{
				Fecha = dr.GetDateTime(0);
				this.cboFecha.Items.Add(Fecha.ToString("dd/MM/yyyy"));
			}
			dr.Close();

			this.cboFecha.Text = Fecha.ToString("dd/MM/yyyy");
			
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInforme_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{				
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales para el Cobro
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["Cobrar"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Cobrar"].MaskInput = stInput;			
			#endregion Decimales para el Cobro
		}

		private void uGridInforme_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad Justifica
			if (e.Cell.Column.ToString() == "CantidadJustificado")
			{
				if ((int)e.Cell.Row.Cells["idEstadoJust"].Value == 3 && (int)e.Cell.Row.Cells["CantidadJustificado"].Value > 0)
				{
					e.Cell.Value = 0;
					MessageBox.Show("Seleccione Un Estado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else if ((int)e.Cell.Row.Cells["CantidadJustificado"].Value > Math.Abs((int)e.Cell.Row.Cells["ConteoExistencia"].Value))
				{
					e.Cell.Value = 0;

					string sMensaje = "";
					if ((int)e.Cell.Row.Cells["idEstadoJust"].Value == 5) sMensaje = " A JUSTIFICAR ";
					else if ((int)e.Cell.Row.Cells["idEstadoJust"].Value == 6) sMensaje = " A COBRAR ";
					else if ((int)e.Cell.Row.Cells["idEstadoJust"].Value == 7) sMensaje = " DE LA PERDIDA ";

					MessageBox.Show(string.Format("La Cantidad '{0}' No puede ser Mayor a la Diferencia", sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else 
				{
					if ((int)e.Cell.Row.Cells["idEstadoJust"].Value == 6)
					{
						decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
						decimal dPrecio = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Precio3 From Articulo Where idArticulo = {0}", (int)e.Cell.Row.Cells["idArticulo"].Value));
						int iCantidad = (int)e.Cell.Row.Cells["CantidadJustificado"].Value;
						decimal dCobro = Convert.ToDecimal((dPrecio * iCantidad) * dIVA);
						e.Cell.Row.Cells["Cobrar"].Value = dCobro;						
					}
					else 
					{
						e.Cell.Row.Cells["Cobrar"].Value = 0.00;
					}

					string sSQLAC = string.Format("Update LogAjusteCantidades Set CantidadJustificado = {0} Where idLogAjusteCantidades = {1}", (int)e.Cell.Row.Cells["CantidadJustificado"].Value, (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAC, true);
				}
			}
			#endregion Cantidad Justifica

			#region Estado
			if (e.Cell.Column.ToString() == "idEstadoJust")
			{	
				string sSQLAE = string.Format("Update LogAjusteCantidades Set idEstadoJust = {0} Where idLogAjusteCantidades = {1}", (int)e.Cell.Row.Cells["idEstadoJust"].Value, (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAE, true);
			}
			#endregion Estado

			#region Notas
			if (e.Cell.Column.ToString() == "NotasCierre")
			{
				string sSQLAN = string.Format("Update LogAjusteCantidades Set NotasCierre = '{0}' Where idLogAjusteCantidades = {1}", e.Cell.Row.Cells["NotasCierre"].Value.ToString(), (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAN, true);
			}
			#endregion Notas

			#region Cobrar
			if (e.Cell.Column.ToString() == "Cobrar")
			{
				if ((int)e.Cell.Row.Cells["idEstadoJust"].Value != 6 && (decimal)e.Cell.Row.Cells["Cobrar"].Value > 0)
				{
					e.Cell.Value = 0.00m;
					MessageBox.Show("Seleccione Estado COBRADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else
				{
					string sSQLACobro = string.Format("Update LogAjusteCantidades Set Cobrar = {0} Where idLogAjusteCantidades = {1}", (decimal)e.Cell.Row.Cells["Cobrar"].Value, (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLACobro, true);
				}
			}
			#endregion Cobrar
		}

		private void uGridInforme_DoubleClick(object sender, System.EventArgs e)
		{
			DateTime dFecha = DateTime.Parse(this.cboFecha.Text);
			using (frmConteoPersonal Personal = new frmConteoPersonal((int)this.uGridInforme.ActiveRow.Cells["idLogAjusteCantidades"].Value, 2,dFecha,(int)this.cmbBodega.Value,1))
			{				
				if (DialogResult.OK == Personal.ShowDialog())
				{
					int iParticipantes = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From LogAjusteCantidadesPersonal Where idLogAjusteCantidades = {0}", (int)this.uGridInforme.ActiveRow.Cells["idLogAjusteCantidades"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set Participantes = {0} Where idLogAjusteCantidades = {1}", iParticipantes, (int)this.uGridInforme.ActiveRow.Cells["idLogAjusteCantidades"].Value));
				}
			}
		}

		private void uGridInforme_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridInforme.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							//							if (uGridInforme.ActiveCell.DroppedDown) return;

							uGridInforme.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridInforme.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							uGridInforme.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							uGridInforme.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridInforme.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							uGridInforme.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (uGridInforme.ActiveCell.DroppedDown) return;
			
							uGridInforme.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridInforme.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							uGridInforme.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							uGridInforme.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridInforme.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGridInforme.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(uGridInforme.ActiveCell.DroppedDown == false) uGridInforme.ActiveCell.DroppedDown = true;
							//						uGridInforme.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void cboFecha_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

