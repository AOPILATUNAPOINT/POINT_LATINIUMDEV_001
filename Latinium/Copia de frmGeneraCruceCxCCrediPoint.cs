using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGeneraCruceCxCCrediPoint.
	/// </summary>
	public class frmGeneraCruceCxCCrediPoint : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		public DevExpress.XtraEditors.SimpleButton btVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public DevExpress.XtraEditors.SimpleButton btnContabilizar;
		public DevExpress.XtraEditors.SimpleButton btnParametros;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLiquidado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAcreditado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFondoGarantia;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsiento;
		public DevExpress.XtraEditors.SimpleButton btnImprimir;
		public DevExpress.XtraEditors.SimpleButton btnAsiento;
		private System.Windows.Forms.Label lblEstado;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmGeneraCruceCxCCrediPoint()
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

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCuotas");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fondo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ACreditado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalCuotas", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalCuotas", 8, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 9, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Fondo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Fondo", 10, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ACreditado", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ACreditado", 11, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fondo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ACreditado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGeneraCruceCxCCrediPoint));
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.btnContabilizar = new DevExpress.XtraEditors.SimpleButton();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnParametros = new DevExpress.XtraEditors.SimpleButton();
			this.txtLiquidado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAcreditado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFondoGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtidAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btnAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLiquidado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAcreditado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).BeginInit();
			this.SuspendLayout();
			// 
			// txtLote
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance1;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(48, 8);
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(96, 22);
			this.txtLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLote.TabIndex = 0;
			this.txtLote.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLote_KeyDown);
			this.txtLote.Click += new System.EventHandler(this.txtLote_Click);
			this.txtLote.Validated += new System.EventHandler(this.txtLote_Validated);
			this.txtLote.ValueChanged += new System.EventHandler(this.txtLote_ValueChanged);
			this.txtLote.Enter += new System.EventHandler(this.txtLote_Enter);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(26, 16);
			this.label3.TabIndex = 212;
			this.label3.Text = "Lote";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance2;
			this.uGridLotes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 32;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 59;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Local";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "N. Operación";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 107;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Documento";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 107;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 297;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Liquido";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 99;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 99;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "F. Garantía";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 99;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance6;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Acreditado";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 107;
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
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance10;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance12;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance13.BackColor = System.Drawing.SystemColors.Control;
			appearance13.BackColor2 = System.Drawing.SystemColors.Control;
			appearance13.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance13.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance15;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 88);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(1018, 440);
			this.uGridLotes.TabIndex = 214;
			this.uGridLotes.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridLotes_AfterCellUpdate);
			this.uGridLotes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLotes_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn12.DefaultValue = 0;
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
			// btVer
			// 
			this.btVer.Location = new System.Drawing.Point(160, 8);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(72, 22);
			this.btVer.TabIndex = 1;
			this.btVer.Text = "&Ver";
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.Location = new System.Drawing.Point(928, 50);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(96, 22);
			this.btnContabilizar.TabIndex = 2;
			this.btnContabilizar.Text = "&Contabilizar";
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 51);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 620;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1034, 8);
			this.groupBox1.TabIndex = 621;
			this.groupBox1.TabStop = false;
			// 
			// btnParametros
			// 
			this.btnParametros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnParametros.Location = new System.Drawing.Point(928, 8);
			this.btnParametros.Name = "btnParametros";
			this.btnParametros.Size = new System.Drawing.Size(96, 22);
			this.btnParametros.TabIndex = 3;
			this.btnParametros.Text = "P&arámetros";
			this.btnParametros.Click += new System.EventHandler(this.btnParametros_Click);
			// 
			// txtLiquidado
			// 
			this.txtLiquidado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.ForeColor = System.Drawing.Color.Blue;
			appearance19.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtLiquidado.Appearance = appearance19;
			this.txtLiquidado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLiquidado.Enabled = false;
			this.txtLiquidado.FormatString = "#,##0.00";
			this.txtLiquidado.Location = new System.Drawing.Point(610, 504);
			this.txtLiquidado.Name = "txtLiquidado";
			this.txtLiquidado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtLiquidado.PromptChar = ' ';
			this.txtLiquidado.Size = new System.Drawing.Size(100, 22);
			this.txtLiquidado.TabIndex = 622;
			this.txtLiquidado.Visible = false;
			// 
			// txtComision
			// 
			this.txtComision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColor = System.Drawing.Color.Blue;
			appearance20.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtComision.Appearance = appearance20;
			this.txtComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComision.Enabled = false;
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(714, 504);
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(100, 22);
			this.txtComision.TabIndex = 623;
			this.txtComision.Visible = false;
			// 
			// txtAcreditado
			// 
			this.txtAcreditado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance21.ForeColor = System.Drawing.Color.Blue;
			appearance21.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtAcreditado.Appearance = appearance21;
			this.txtAcreditado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAcreditado.Enabled = false;
			this.txtAcreditado.FormatString = "#,##0.00";
			this.txtAcreditado.Location = new System.Drawing.Point(922, 504);
			this.txtAcreditado.Name = "txtAcreditado";
			this.txtAcreditado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAcreditado.PromptChar = ' ';
			this.txtAcreditado.Size = new System.Drawing.Size(100, 22);
			this.txtAcreditado.TabIndex = 625;
			this.txtAcreditado.Visible = false;
			// 
			// txtFondoGarantia
			// 
			this.txtFondoGarantia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance22.ForeColor = System.Drawing.Color.Blue;
			appearance22.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtFondoGarantia.Appearance = appearance22;
			this.txtFondoGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFondoGarantia.Enabled = false;
			this.txtFondoGarantia.FormatString = "#,##0.00";
			this.txtFondoGarantia.Location = new System.Drawing.Point(818, 504);
			this.txtFondoGarantia.Name = "txtFondoGarantia";
			this.txtFondoGarantia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFondoGarantia.PromptChar = ' ';
			this.txtFondoGarantia.Size = new System.Drawing.Size(100, 22);
			this.txtFondoGarantia.TabIndex = 624;
			this.txtFondoGarantia.Visible = false;
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcel.Location = new System.Drawing.Point(840, 8);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(80, 22);
			this.btExcel.TabIndex = 627;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(744, 53);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 629;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(800, 51);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(120, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 628;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			this.txtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFecha_KeyDown);
			this.txtFecha.Validated += new System.EventHandler(this.txtFecha_Validated);
			this.txtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.txtFecha_BeforeDropDown);
			// 
			// txtidAsiento
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidAsiento.Appearance = appearance23;
			this.txtidAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtidAsiento.Enabled = false;
			this.txtidAsiento.FormatString = "";
			this.txtidAsiento.Location = new System.Drawing.Point(504, 504);
			this.txtidAsiento.Name = "txtidAsiento";
			this.txtidAsiento.PromptChar = ' ';
			this.txtidAsiento.Size = new System.Drawing.Size(96, 22);
			this.txtidAsiento.TabIndex = 630;
			this.txtidAsiento.Visible = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Location = new System.Drawing.Point(632, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(96, 22);
			this.btnImprimir.TabIndex = 631;
			this.btnImprimir.Text = "&Imprimir Asiento";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsiento.Location = new System.Drawing.Point(736, 8);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(96, 22);
			this.btnAsiento.TabIndex = 632;
			this.btnAsiento.Text = "&Ver Asiento";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(416, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 633;
			// 
			// frmGeneraCruceCxCCrediPoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1034, 534);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btnParametros);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.uGridLotes);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtidAsiento);
			this.Controls.Add(this.txtAcreditado);
			this.Controls.Add(this.txtFondoGarantia);
			this.Controls.Add(this.txtComision);
			this.Controls.Add(this.txtLiquidado);
			this.Name = "frmGeneraCruceCxCCrediPoint";
			this.Text = "Liquidación Cartera Solidario";
			this.Load += new System.EventHandler(this.frmGeneraCruceCxCCrediPoint_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLiquidado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAcreditado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec ConsultaLotesCrediPoint {0}", (int)this.txtLote.Value);
			this.uGridLotes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			int idAsiento = 0;
			this.txtidAsiento.Value = 0;
			idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idAsiento Int = 0 If Exists(Select ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idOrigenAsiento = 18 And idTipoAsiento = 1 And Borrar = 0) Select @idAsiento = ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idTipoAsiento = 1 And Borrar = 0 Select @idAsiento", this.txtLote.Value.ToString()));
			if (idAsiento > 0)
			{
				this.txtidAsiento.Value = idAsiento;
				this.lblEstado.Text = "CONTABILIZADO";
				//	this.btnContabilizar.Enabled = false;
			}
			else 
			{
				this.lblEstado.Text = "PENDIENTE";
				this.btnContabilizar.Enabled = true;
			}

			Sumatorias();

			ContSel();
		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Consulta();
		}

		private void frmGeneraCruceCxCCrediPoint_Load(object sender, System.EventArgs e)
		{			
			miAcceso = new Acceso(cdsSeteoF, "08320105");

			if (!Funcion.Permiso("959", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso REPORTE LIQUIDACION CARTERA SOLIDARIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}
			this.txtLote.Value = Funcion.iEscalarSQL(cdsSeteoF, "Exec RetornaUltimoLoteSolidario");
			this.txtFecha.Value = DateTime.Today;

			this.btVer_Click(sender, e);
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaPeriodo((DateTime)this.txtFecha.Value, cdsSeteoF, "Bancos"))
			{
				DateTime dtFechaVP = (DateTime)this.txtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Asientos en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFecha.Focus();
				return;
			}
			
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				return;
			}
			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Documentos con este Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			this.Cursor = Cursors.WaitCursor;
			
			Funcion.TiempoInicia();
			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;
						
					string sSQL = string.Format("Exec AsientoLoteCrediPoint {0}, {1}, '{2}', {3}, {4}, {5}", 
						(int)this.txtidAsiento.Value, (int)this.txtLote.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), 
						Convert.ToDecimal(this.txtComision.Value), Convert.ToDecimal(this.txtFondoGarantia.Value), Convert.ToDecimal(this.txtAcreditado.Value));					
					oCmdActualiza.CommandText = sSQL;
					this.txtidAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

					int iCont = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{
						if ((bool)dr.Cells["Sel"].Value)
						{							
							oCmdActualiza.CommandText = string.Format("Update Compra Set idAsientoLoteSolidario = {0}, FechaVenta = '{2}', ValorAcreditado = {3}, Comision = {4}, FondoDeGarantia = {5}, ComisionLote = ({4} + {5}), Bloquear = 1, Confirmado = 1 Where idCompra = {1}", 
								(int)this.txtidAsiento.Value, (int)dr.Cells["idCompra"].Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), 
								Convert.ToDecimal(dr.Cells["ACreditado"].Value), Convert.ToDecimal(dr.Cells["Comision"].Value), Convert.ToDecimal(dr.Cells["Fondo"].Value));						
							oCmdActualiza.ExecuteNonQuery();

							oCmdActualiza.CommandText = string.Format("Exec ActualizaSaldoFacturas {0}, 1, 9", (int)dr.Cells["idCompra"].Value);
							oCmdActualiza.ExecuteNonQuery();							
					
							iCont++;
						}
					}

					oTransaction.Commit();	

					MessageBox.Show(string.Format("Se Contabilizó {1} documentos del Lote {0} " + Funcion.Tiempo("EN").ToString(), (int)this.txtLote.Value, iCont.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.lblEstado.Text = "CONTABILIZADO";

					#region Impresion
					string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();
				
					Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("Asiento");			
					miRepor.Show();	
					#endregion Impresion

					this.btVer_Click(sender, e);

					this.Cursor = Cursors.Default;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

					oConexion.Close();					

					this.Cursor = Cursors.Default;
				}
				finally
				{
					oConexion.Close();

					this.Cursor = Cursors.Default;
				}
			}			
		}

		private void btnParametros_Click(object sender, System.EventArgs e)
		{
			using (frmParametrosLotes Parametros = new frmParametrosLotes())
			{
				if (DialogResult.OK == Parametros.ShowDialog())
				{
					string sSQL = string.Format("Update ParametrosCruceCP Set Comision = {0}, FondoDeGarantia = {1}, idCuenta = {2}", 
						(double)Parametros.txtComision.Value, (double)Parametros.txtFondoGarantia.Value, (double)Parametros.cmbCuentas.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					this.btVer_Click(sender, e);
				}
			}
		}

		private void txtLote_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void Sumatorias()
		{
			this.txtLiquidado.Value = 0;
			this.txtComision.Value = 0;
			this.txtFondoGarantia.Value = 0;
			this.txtAcreditado.Value = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{	
					this.txtLiquidado.Value = Convert.ToDecimal(this.txtLiquidado.Value) + Convert.ToDecimal(dr.Cells["TotalCuotas"].Value);
					this.txtComision.Value = Convert.ToDecimal(this.txtComision.Value) + Convert.ToDecimal(dr.Cells["Comision"].Value);
					this.txtFondoGarantia.Value = Convert.ToDecimal(this.txtFondoGarantia.Value) + Convert.ToDecimal(dr.Cells["Fondo"].Value);
					this.txtAcreditado.Value = Convert.ToDecimal(this.txtAcreditado.Value) + Convert.ToDecimal(dr.Cells["ACreditado"].Value);
				}
			}
		}

		private void ContSel()
		{
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value) iCont++;
			}

			this.lblContador.Text = this.uGridLotes.Rows.Count + " DOCUMENTOS ENCONTRADOS - " + iCont.ToString() + " SELECCIONADOS";
		}

		private void uGridLotes_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				Sumatorias();

				ContSel();
			}
		}

		private void uGridLotes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			FuncionesProcedimientos.ExportaExcel(this.uGridLotes);
		}

		private void txtLote_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.btVer_Click(sender, e);
				this.txtFecha.Focus();
			}
		}

		private void txtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.txtFecha.Value == System.DBNull.Value) this.txtFecha.Value = DateTime.Today;
		}

		private void txtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnContabilizar.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Documentos con este Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				#region Impresion
				string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();
				
				Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Asiento");			
				miRepor.Show();	
				#endregion Impresion
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Documentos con este Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.txtidAsiento.Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;
			}
		}

		private void txtLote_Validated(object sender, System.EventArgs e)
		{
			this.btVer_Click(sender, e);
		}

		private void txtLote_Click(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll();
		}

		private void txtLote_Enter(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll();
		}

		private void txtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}

