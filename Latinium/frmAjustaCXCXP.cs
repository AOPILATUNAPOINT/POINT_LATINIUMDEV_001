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
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAjustaCXCXP.
	/// </summary>
	public class frmAjustaCXCXP : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		public DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label2;
		public DevExpress.XtraEditors.SimpleButton btnProcesar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		public DevExpress.XtraEditors.SimpleButton btnImprimir;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivar;
		public DevExpress.XtraEditors.SimpleButton btnActualizarTabla;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAjustaCXCXP()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumDocAntiguo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo1");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 10, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 11, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo1", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo1", 12, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumDocAntiguo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo1");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAjustaCXCXP));
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnProcesar = new DevExpress.XtraEditors.SimpleButton();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.chkActivar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnActualizarTabla = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
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
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 96;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.Header.Caption = "Numero EM";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 96;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Doc";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 65;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 52;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 110;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 250;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance2;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 85;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance3;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 85;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 85;
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
																										 ultraGridColumn13});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance6;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance7;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(962, 256);
			this.ultraGrid1.TabIndex = 3;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn13.DataType = typeof(System.Double);
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 633;
			this.label1.Text = "Fecha";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(55, 16);
			this.label3.TabIndex = 631;
			this.label3.Text = "Proveedor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance13;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(64, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 628;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(336, 72);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 22);
			this.btBuscar.TabIndex = 2;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// txtBuscar
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance14;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(88, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Location = new System.Drawing.Point(552, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 635;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 16);
			this.label2.TabIndex = 638;
			this.label2.Text = "Documento";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance15;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(88, 72);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(200, 22);
			this.txtNumero.TabIndex = 1;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Location = new System.Drawing.Point(424, 72);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(80, 22);
			this.btnProcesar.TabIndex = 5;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(904, 8);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 640;
			this.txtNumIdCliente.Visible = false;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblCliente.Location = new System.Drawing.Point(312, 40);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(0, 23);
			this.lblCliente.TabIndex = 641;
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.label4.Location = new System.Drawing.Point(224, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(232, 20);
			this.label4.TabIndex = 642;
			this.label4.Text = "F3 PARA BUSCAR PROVEEDORES";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(978, 8);
			this.groupBox1.TabIndex = 643;
			this.groupBox1.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Location = new System.Drawing.Point(512, 72);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 22);
			this.btnImprimir.TabIndex = 644;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// chkActivar
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivar.Appearance = appearance16;
			this.chkActivar.BackColor = System.Drawing.Color.Transparent;
			this.chkActivar.Location = new System.Drawing.Point(776, 40);
			this.chkActivar.Name = "chkActivar";
			this.chkActivar.Size = new System.Drawing.Size(64, 22);
			this.chkActivar.TabIndex = 645;
			this.chkActivar.Text = "Activar";
			this.chkActivar.CheckedChanged += new System.EventHandler(this.chkActivar_CheckedChanged);
			// 
			// btnActualizarTabla
			// 
			this.btnActualizarTabla.Location = new System.Drawing.Point(632, 72);
			this.btnActualizarTabla.Name = "btnActualizarTabla";
			this.btnActualizarTabla.Size = new System.Drawing.Size(80, 22);
			this.btnActualizarTabla.TabIndex = 646;
			this.btnActualizarTabla.Text = "&Arreglar Tabla";
			this.btnActualizarTabla.Click += new System.EventHandler(this.btnActualizarTabla_Click);
			// 
			// frmAjustaCXCXP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(978, 390);
			this.Controls.Add(this.btnActualizarTabla);
			this.Controls.Add(this.chkActivar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "frmAjustaCXCXP";
			this.Text = "Ajusta CXC CXP";
			this.Load += new System.EventHandler(this.frmAjustaCXCXP_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
			
			this.ultraGrid1.DataSource = ultraDataSource1;
			
//			if ((int)this.txtNumIdCliente.Value == 0)
//			{
//				MessageBox.Show("Ingrese un Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//			}
//			else if (this.txtBuscar.Text.ToString().Length == 0 && this.txtNumero.Text.ToString().Length > 0)
//			{
//				MessageBox.Show("Ingrese el Proveedor para Buscar Documentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//			}
//			else
//			{
				string sSQL = string.Format("Exec ReporteAjustaCXP '{0}', '{1}', '{2}'", 
					this.txtBuscar.Text.ToString(), this.txtNumero.Text.ToString(), dtHasta.ToString("yyyyMMdd"));

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.lblContador.Text = this.ultraGrid1.Rows.Count + " DOCUMENTOS";
//			}
		}

		private void frmAjustaCXCXP_Load(object sender, System.EventArgs e)
		{
			//string sHasta = "31/12/2014";
			//DateTime dtHasta = DateTime.DaysInMonth sHasta.ToString("dd/MM/yyyy");// DateTime.Parse(sHasta, "dd/MM/yyyy");

			this.dtHasta.Value = "31/12/2014";			
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtBuscar.Text.ToString(), 4, 1, 1, 1, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtBuscar.Text = miBusqueda.Ruc;
						
						bool bActivar = false;
						this.chkActivar.Checked = false;
						
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idCliente, 0), ISNULL(Nombre, ''), Sel From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtBuscar.Text.ToString()), true);
						dr.Read();
						if (dr.HasRows)
						{
							this.txtNumIdCliente.Value = dr.GetInt32(0);
							this.lblCliente.Text = dr.GetString(1);
							bActivar = dr.GetBoolean(2);							
						}
						dr.Close();

						this.chkActivar.Checked = bActivar;

						this.btBuscar_Click(sender, e);

						this.txtNumero.Focus();
					}
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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
						
			#region Decimales para el Saldo 1
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["Saldo1"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Saldo1"].MaskInput = stInput;			
			#endregion Decimales para el Saldo 1
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				if ((int)e.Cell.Row.Cells["idTipoFactura"].Value != 25)
				{
					string sSQL = string.Format("Exec ActualizaSelCompraCXCXP {0}, {1}", (bool)e.Cell.Row.Cells["Sel"].Value, (int)e.Cell.Row.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
				}
				else
				{
					string sSQL = string.Format("Exec ActualizaSelAnticipo {0}, {1}", (bool)e.Cell.Row.Cells["Sel"].Value, (int)e.Cell.Row.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
				}

			}

			if (e.Cell.Column.ToString() == "NumDocAntiguo")
			{
				string sSQLNEM = string.Format("Update Compra Set NumDocAntiguo = '{0}' Where idCompra = {1}", e.Cell.Row.Cells["NumDocAntiguo"].Value.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLNEM, true);	
			}

			if (e.Cell.Column.ToString() == "Saldo1")
			{
				string sSQLSaldo = string.Format("Update Compra Set Saldo1 = {0} Where idCompra = {1}", (double)e.Cell.Row.Cells["Saldo1"].Value, (int)e.Cell.Row.Cells["idCompra"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);
			}	
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;

							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;
			
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
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

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btBuscar_Click(sender, e);
		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{	
			if ((int)this.txtNumIdCliente.Value == 0)
			{
				MessageBox.Show("Ingrese un Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.txtBuscar.Text.ToString().Length == 0 && this.txtNumero.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Proveedor para Buscar Documentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Datos en la Lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else 
			if (DialogResult.Yes ==	MessageBox.Show("Antes de Generar los Ajustes : \n\n * Revise que los Documentos y Valores a Ajustar sean los correctos \n\n * Revise que los Documentos Pendientes de Pago se encuentren Marcados \n\n\n ¿Esta Seguro de Procesar? ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					this.Cursor = Cursors.WaitCursor;

					#region Facturas
					if ((int)dr.Cells["idTipoFactura"].Value == 4)
					{
						if (!(bool)dr.Cells["Sel"].Value)
						{
							string sSQL = string.Format("Exec CreaNotaDeCreditoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, (double)dr.Cells["Saldo"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
						else if ((bool)dr.Cells["Sel"].Value)
						{

							double dValorNC = 0;
							double dSaldoActual = (double)dr.Cells["Saldo"].Value;
							double dSaldo1 = (double)dr.Cells["Saldo1"].Value;

							if (dSaldo1 > 0)
							{
								dValorNC = dSaldoActual - dSaldo1;

								if (dValorNC > 0)
								{
									string sSQL = string.Format("Exec CreaNotaDeCreditoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, dValorNC);
									Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
								}
							}

							this.Cursor = Cursors.Default;
						}
					}
					#endregion Facturas

					#region Notas de Credito
					if ((int)dr.Cells["idTipoFactura"].Value == 6)
					{
						if (!(bool)dr.Cells["Sel"].Value)
						{
							string sSQL = string.Format("Exec CreaNotaDeDebitoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, Math.Abs((double)dr.Cells["Saldo"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
						else if ((bool)dr.Cells["Sel"].Value)
						{

							double dValorNC = 0;
							double dSaldoActual = Math.Abs((double)dr.Cells["Saldo"].Value);
							double dSaldo1 = (double)dr.Cells["Saldo1"].Value;

							if (dSaldo1 > 0)
							{
								dValorNC = dSaldoActual - dSaldo1;

								if (dValorNC > 0)
								{
									string sSQL = string.Format("Exec CreaNotaDeDebitoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, dValorNC);
									Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
								}
							}

							this.Cursor = Cursors.Default;
						}
					}
					#endregion Notas de Credito

					#region Anticipo
					/*
					if ((int)dr.Cells["idTipoFactura"].Value == 25)
					{
						if (!(bool)dr.Cells["Sel"].Value)
						{
							string sSQL = string.Format("Exec CreaNotaDeCreditoAnticipoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, Math.Abs((double)dr.Cells["Saldo"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
						else if ((bool)dr.Cells["Sel"].Value)
						{

							double dValorNC = 0;
							double dSaldoActual = Math.Abs((double)dr.Cells["Saldo"].Value);
							double dSaldo1 = (double)dr.Cells["Saldo1"].Value;

							if (dSaldo1 > 0)
							{
								dValorNC = dSaldoActual - dSaldo1;

								if (dValorNC > 0)
								{
									string sSQL = string.Format("Exec CreaNotaDeCreditoAnticipoAjusteCXP {0}, {1}", (int)dr.Cells["idCompra"].Value, dValorNC);
									Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
								}
							}					
						}
					}
					*/
					#endregion Anticipo
					this.Cursor = Cursors.Default;
				}

				MessageBox.Show("Proceso Finalizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btBuscar_Click(sender, e);
			}
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfCapital = new ParameterField ();
			ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
			pfCapital.ParameterFieldName = "@Filtro";
			ValCapital.Value = "";
			pfCapital.CurrentValues.Add (ValCapital);
			paramFields.Add (pfCapital);
				
			ParameterField pfCuotas = new ParameterField ();
			ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
			pfCuotas.ParameterFieldName = "@Numero";
			ValCuotas.Value = "";
			pfCuotas.CurrentValues.Add (ValCuotas);
			paramFields.Add (pfCuotas);
				
			ParameterField pfFecha = new ParameterField ();
			ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
			pfFecha.ParameterFieldName = "@Fecha";
			ValFecha.Value = dtHasta;
			pfFecha.CurrentValues.Add (ValFecha);
			paramFields.Add (pfFecha);
				
			Reporte miRepor = new Reporte("TemporalCXC.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	
		}

		private void chkActivar_CheckedChanged(object sender, System.EventArgs e)
		{
//			string sSQL = string.Format("Exec ClienteActualizaSel {0}, {1}", (int)this.txtNumIdCliente.Value, this.chkActivar.Checked);
//			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
		}

		private void btnActualizarTabla_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{			
					ultraGridExcelExporter1.Export(this.ultraGrid1, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}
	}
}

