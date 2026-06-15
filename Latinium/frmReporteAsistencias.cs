using System;
using System.Drawing;
using System.Collections;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAsistencias.
	/// </summary>
	public class frmReporteAsistencias : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnGenerarPago;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoPagoAsistencia;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;

		private int idGrupo;

		public frmReporteAsistencias()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAsistencias));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Producto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorUnit");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 6, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 8, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Producto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorUnit");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGenerarPago = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstadoPagoAsistencia = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoPagoAsistencia)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 212;
			this.label1.Text = "Bodega";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance1;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Width = 60;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 38);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 21);
			this.cmbBodega.TabIndex = 211;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 297;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label6.Click += new System.EventHandler(this.label6_Click);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 296;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 16);
			this.label2.TabIndex = 299;
			this.label2.Text = "Vendedor";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// cmbVendedor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance3;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 224;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(72, 70);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(256, 21);
			this.cmbVendedor.TabIndex = 300;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(368, 68);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 301;
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance4.BackColor = System.Drawing.Color.White;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Appearance = appearance4;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 82;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 133;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 311;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 182;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Tipo Precio";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 122;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Header.Caption = "Valor Unit";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance7;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 70;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand3.Header.Caption = "COMISIONES";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
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
			this.ultraGrid2.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1136, 264);
			this.ultraGrid2.TabIndex = 302;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// btnGenerarPago
			// 
			this.btnGenerarPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarPago.CausesValidation = false;
			this.btnGenerarPago.Enabled = false;
			this.btnGenerarPago.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarPago.Image")));
			this.btnGenerarPago.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerarPago.Location = new System.Drawing.Point(1048, 7);
			this.btnGenerarPago.Name = "btnGenerarPago";
			this.btnGenerarPago.Size = new System.Drawing.Size(95, 23);
			this.btnGenerarPago.TabIndex = 639;
			this.btnGenerarPago.Text = "&Generar Pago";
			this.btnGenerarPago.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerarPago.Click += new System.EventHandler(this.btnGenerarPago_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(208, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 641;
			this.label3.Text = "Hasta";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// dtHasta
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance16;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(264, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 640;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// optEstadoPagoAsistencia
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoPagoAsistencia.Appearance = appearance17;
			this.optEstadoPagoAsistencia.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoPagoAsistencia.CheckedIndex = 0;
			this.optEstadoPagoAsistencia.ItemAppearance = appearance18;
			this.optEstadoPagoAsistencia.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Pendientes";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Pagados";
			this.optEstadoPagoAsistencia.Items.Add(valueListItem1);
			this.optEstadoPagoAsistencia.Items.Add(valueListItem2);
			this.optEstadoPagoAsistencia.ItemSpacingVertical = 10;
			this.optEstadoPagoAsistencia.Location = new System.Drawing.Point(472, 5);
			this.optEstadoPagoAsistencia.Name = "optEstadoPagoAsistencia";
			this.optEstadoPagoAsistencia.Size = new System.Drawing.Size(176, 26);
			this.optEstadoPagoAsistencia.TabIndex = 1063;
			this.optEstadoPagoAsistencia.Text = "Pendientes";
			this.optEstadoPagoAsistencia.ValueChanged += new System.EventHandler(this.optEstadoPagoAsistencia_ValueChanged);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1048, 37);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(96, 23);
			this.btnImprimir.TabIndex = 1064;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1048, 69);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(96, 23);
			this.btnExportar.TabIndex = 1082;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1152, 8);
			this.groupBox1.TabIndex = 1083;
			this.groupBox1.TabStop = false;
			// 
			// frmReporteAsistencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1152, 390);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.optEstadoPagoAsistencia);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btnGenerarPago);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbBodega);
			this.Name = "frmReporteAsistencias";
			this.Text = "Reporte de Intangibles ";
			this.Load += new System.EventHandler(this.frmReporteAsistencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoPagoAsistencia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmReporteAsistencias_Load(object sender, System.EventArgs e)
		{
			DateTime hoy = DateTime.Today;
			DateTime primerDiaDeJunio = new DateTime(hoy.Year, 6, 1);
			DateTime primerDiaDelMes = new DateTime(hoy.Year, hoy.Month, 1);

			this.dtDesde.CalendarInfo.MinDate = primerDiaDeJunio;
			this.dtDesde.CalendarInfo.MaxDate = hoy;

			this.dtHasta.CalendarInfo.MinDate = primerDiaDeJunio;
			this.dtHasta.CalendarInfo.MaxDate = hoy;

			this.dtDesde.Value = primerDiaDelMes;		
			this.dtHasta.Value = hoy;		

			string sSQL = string.Format("Exec ListaBodegas {0}, 43", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);


			string exportarSQL = string.Format("select idGrupo from Usuario where idUsuario = {0}", MenuLatinium.IdUsuario);
			idGrupo = Funcion.iEscalarSQL(cdsSeteoF, exportarSQL);	
			
			if (idGrupo == 1 || idGrupo == 15 || idGrupo == 28) 
			{
				this.btnExportar.Enabled = true;
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{

			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
		if (this.cmbBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una bodega válida de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				
				return;
			}

			int idBodega = Convert.ToInt32(this.cmbBodega.Value);
			if (idBodega == 0) 
			{
//				this.btnGenerarPago.Enabled = false;
				this.cmbVendedor.DataSource = null;
			} 
			else if((int)this.optEstadoPagoAsistencia.Value != 2)
			{
//				this.btnVer.Enabled = true;
//				this.btnGenerarPago.Enabled = true;
				string sSQL = string.Format("Exec ListaUsuarios {0}", idBodega);
				this.cmbVendedor.Value = 0;
				this.cmbVendedor.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			// Validación básica de controles
			if (this.cmbBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una bodega válida de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.btnVer.Enabled = false;
				return;
			}
			if (this.cmbVendedor.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un vendedor válido antes de continuar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			if (this.dtDesde.Value == System.DBNull.Value || this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha válida antes de continuar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			string sDesde = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");
			string sHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");
			int idBodega = Convert.ToInt32(this.cmbBodega.Value);
			int idVendedor = Convert.ToInt32(this.cmbVendedor.Value);
			int iEstadoPagoAsistencia = Convert.ToInt32(this.optEstadoPagoAsistencia.Value) ;

			string sSQL = string.Format("Exec SP_ReporteAsistencias '{0}', '{1}', {2}, {3}, {4}", sDesde, sHasta, idBodega, idVendedor, iEstadoPagoAsistencia);

			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid2);

			// Validar si el grid tiene filas para permitir generar el pago
			if (this.ultraGrid2.Rows.Count > 0 && this.cmbBodega.Value != null && (int)this.cmbBodega.Value != 0 && (int)this.optEstadoPagoAsistencia.Value != 2)
			{
				this.btnGenerarPago.Enabled = true;
				this.btnImprimir.Enabled = false;
			}
			else if (this.ultraGrid2.Rows.Count > 0 && (int)this.optEstadoPagoAsistencia.Value == 2)
			{
				this.btnGenerarPago.Enabled = false;
				this.btnImprimir.Enabled = true;
			}
			else
			{
				this.btnGenerarPago.Enabled = false;
				this.btnImprimir.Enabled = false;
			}
			
		}

		private void btnGenerarPago_Click(object sender, System.EventArgs e)
		{
			// Validación básica de controles
			if (this.cmbBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una bodega válida de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.btnGenerarPago.Enabled = false;
				return;
			}
			if (DialogResult.Yes == MessageBox.Show(
				"¿Está seguro que desea generar el pago para la bodega seleccionada?",
				"Confirmación",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2))
			{
				#region Variables

				int idBodega = Convert.ToInt32(this.cmbBodega.Value);
				string sDesde = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");
				string sHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");
				int idVendedor = Convert.ToInt32(this.cmbVendedor.Value);

				int idPagoAsistencias = 0;

				#endregion Variables

				#region Guarda Pago Asistencia, Autorizacion Pagos y Actualizacion de Estados 
			
				string Codigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 99, {0}, 0", idBodega));

				string sSQLGPA= string.Format("Exec SP_GenerarPagoAsistencias '{0}', '{1}', {2}, {3}, '{4}'", sDesde, sHasta, idBodega, idVendedor, Codigo);
				idPagoAsistencias = Funcion.iEscalarSQL(cdsSeteoF, sSQLGPA);	
	
				#endregion Guarda Pago Asistencia, Autorizacion Pagos y Actualizacion de Estados 

				#region Impresion Pago Asistencias

				// Crear parámetros para el reporte
				ParameterFields paramFields = new ParameterFields();
				ParameterField paramFieldIdPagoAsistencias = new ParameterField ();
				ParameterDiscreteValue discreteValIdPagoAsistencias = new ParameterDiscreteValue ();
				paramFieldIdPagoAsistencias.ParameterFieldName = "@idPagoAsistencias";
				discreteValIdPagoAsistencias.Value = idPagoAsistencias; 
				paramFieldIdPagoAsistencias.CurrentValues.Add (discreteValIdPagoAsistencias);
				paramFields.Add (paramFieldIdPagoAsistencias);

				// Cargar y mostrar el reporte
				Reporte reporte = new Reporte("ImpresionPagoAsistencias.rpt", "");
				reporte.MdiParent = this.MdiParent;
				reporte.Titulo("Pago Asistencias");
				reporte.crVista.ParameterFieldInfo = paramFields;
				reporte.Show();

				#endregion Impresion Pago Asistencias

				#region Reseteo del grid

				this.btnGenerarPago.Enabled = false;
				System.Data.DataTable dt = (this.ultraGrid2.DataSource as System.Data.DataTable);
				if (dt != null)
					dt.Clear();

				#endregion Reseteo del grid

				MessageBox.Show("Pago generado correctamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;			

			}
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void optEstadoPagoAsistencia_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optEstadoPagoAsistencia.Value == 2)
			{
				this.btnGenerarPago.Enabled	= false;
			} 
			else 
			{
				this.btnImprimir.Enabled = false;
			}
			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Impresión múltiple de pagos de asistencia

			// Validación básica de controles
			if (this.cmbBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una bodega válida de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.btnVer.Enabled = false;
				return;
			}

			if (this.dtDesde.Value == System.DBNull.Value || this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha válida antes de continuar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			string sDesde = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");
			string sHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");
			int idBodega = Convert.ToInt32(this.cmbBodega.Value);
			int idVendedor = Convert.ToInt32(this.cmbVendedor.Value);

			ArrayList listaIdPagos = new ArrayList();

			string sSQL = string.Format("Exec SP_ListarPagosAsistencias  '{0}', '{1}', {2}, {3}, {4}", sDesde, sHasta, idBodega,idGrupo,idVendedor);

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
	
			if (dr != null)
			{
				while (dr.Read())
				{
					listaIdPagos.Add(Convert.ToInt32(dr["idPagoAsistencias"]));
				}
			}
			dr.Close(); 
			
			// Mostrar los reportes
			for (int i = 0; i < listaIdPagos.Count; i++)
			{
				int idPago = (int)listaIdPagos[i];

				ParameterFields paramFields = new ParameterFields();
				ParameterField paramFieldIdPagoAsistencias = new ParameterField();
				ParameterDiscreteValue discreteValIdPagoAsistencias = new ParameterDiscreteValue();

				paramFieldIdPagoAsistencias.ParameterFieldName = "@idPagoAsistencias";
				discreteValIdPagoAsistencias.Value = idPago;
				paramFieldIdPagoAsistencias.CurrentValues.Add(discreteValIdPagoAsistencias);
				paramFields.Add(paramFieldIdPagoAsistencias);

				Reporte reporte = new Reporte("ImpresionPagoAsistencias.rpt", "");
				reporte.MdiParent = this.MdiParent;
				reporte.Titulo("Pago Asistencias");
				reporte.crVista.ParameterFieldInfo = paramFields;
				reporte.Show();
			}

			#endregion Impresión múltiple
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid2);		
		}
	}
}
