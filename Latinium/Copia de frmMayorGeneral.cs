using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmMayorGeneral.
	/// </summary>
	public class frmMayorGeneral : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btConsultar;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.SimpleButton btnAsiento;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalDebe;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalHaber;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFinal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInicial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private System.Windows.Forms.GroupBox groupBox2;
		private DevExpress.XtraEditors.SimpleButton btnImprimir;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMayorGeneral()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 9, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 10, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 11, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMayorGeneral));
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btConsultar = new DevExpress.XtraEditors.SimpleButton();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtFinal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtInicial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.lblIce = new System.Windows.Forms.Label();
			this.txtTotalDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtTotalHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label5 = new System.Windows.Forms.Label();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalDebe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
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
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 79;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 72;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 226;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Numero";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Referencia";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Nombre";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 226;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 100;
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
			ultraGridBand1.Header.Caption = "";
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnHeaderAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance8;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1232, 544);
			this.ultraGrid1.TabIndex = 7;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
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
			// dtHasta
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance15;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(264, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance16;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 0;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 251;
			this.label2.Text = "Hasta:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 250;
			this.label1.Text = "Desde:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btConsultar
			// 
			this.btConsultar.Location = new System.Drawing.Point(1080, 8);
			this.btConsultar.Name = "btConsultar";
			this.btConsultar.Size = new System.Drawing.Size(80, 21);
			this.btConsultar.TabIndex = 5;
			this.btConsultar.Text = "Generar";
			this.btConsultar.Click += new System.EventHandler(this.btConsultar_Click);
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(600, 11);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(28, 16);
			this.lblComprobante.TabIndex = 257;
			this.lblComprobante.Text = "Final";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(400, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(34, 16);
			this.lblNumero.TabIndex = 256;
			this.lblNumero.Text = "Inicial";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFinal
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFinal.Appearance = appearance17;
			this.txtFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFinal.Enabled = false;
			this.txtFinal.Location = new System.Drawing.Point(648, 8);
			this.txtFinal.Name = "txtFinal";
			this.txtFinal.Size = new System.Drawing.Size(120, 22);
			this.txtFinal.TabIndex = 3;
			this.txtFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFinal_KeyDown);
			// 
			// txtInicial
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInicial.Appearance = appearance18;
			this.txtInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInicial.Location = new System.Drawing.Point(456, 8);
			this.txtInicial.Name = "txtInicial";
			this.txtInicial.Size = new System.Drawing.Size(120, 22);
			this.txtInicial.TabIndex = 2;
			this.txtInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInicial_KeyDown);
			this.txtInicial.TextChanged += new System.EventHandler(this.txtInicial_TextChanged);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Location = new System.Drawing.Point(8, 56);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 21);
			this.btnAsiento.TabIndex = 6;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// lblIce
			// 
			this.lblIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(608, 59);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(59, 16);
			this.lblIce.TabIndex = 318;
			this.lblIce.Text = "Total Debe";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalDebe
			// 
			this.txtTotalDebe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.BackColorDisabled = System.Drawing.Color.White;
			appearance19.BackColorDisabled2 = System.Drawing.Color.White;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalDebe.Appearance = appearance19;
			this.txtTotalDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalDebe.Enabled = false;
			this.txtTotalDebe.FormatString = "#,##0.00";
			this.txtTotalDebe.Location = new System.Drawing.Point(688, 56);
			this.txtTotalDebe.Name = "txtTotalDebe";
			this.txtTotalDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalDebe.PromptChar = ' ';
			this.txtTotalDebe.Size = new System.Drawing.Size(112, 22);
			this.txtTotalDebe.TabIndex = 317;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(824, 59);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(63, 16);
			this.label3.TabIndex = 320;
			this.label3.Text = "Total Haber";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalHaber
			// 
			this.txtTotalHaber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.BackColorDisabled = System.Drawing.Color.White;
			appearance20.BackColorDisabled2 = System.Drawing.Color.White;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalHaber.Appearance = appearance20;
			this.txtTotalHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalHaber.Enabled = false;
			this.txtTotalHaber.FormatString = "#,##0.00";
			this.txtTotalHaber.Location = new System.Drawing.Point(904, 56);
			this.txtTotalHaber.Name = "txtTotalHaber";
			this.txtTotalHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalHaber.PromptChar = ' ';
			this.txtTotalHaber.Size = new System.Drawing.Size(112, 22);
			this.txtTotalHaber.TabIndex = 319;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(1040, 59);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(64, 16);
			this.label4.TabIndex = 322;
			this.label4.Text = "Saldo Total ";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoTotal
			// 
			this.txtSaldoTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance21.BackColorDisabled = System.Drawing.Color.White;
			appearance21.BackColorDisabled2 = System.Drawing.Color.White;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance21;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(1120, 56);
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoTotal.TabIndex = 321;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 88);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1246, 8);
			this.groupBox1.TabIndex = 323;
			this.groupBox1.TabStop = false;
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
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(792, 11);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 325;
			this.label5.Text = "Cuenta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuenta
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuenta.Appearance = appearance22;
			this.txtCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuenta.Location = new System.Drawing.Point(848, 8);
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(216, 22);
			this.txtCuenta.TabIndex = 4;
			this.txtCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuenta_KeyDown);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1248, 8);
			this.groupBox2.TabIndex = 326;
			this.groupBox2.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Location = new System.Drawing.Point(104, 56);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 21);
			this.btnImprimir.TabIndex = 327;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmMayorGeneral
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1246, 660);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtCuenta);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtSaldoTotal);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtTotalHaber);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.txtTotalDebe);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtFinal);
			this.Controls.Add(this.txtInicial);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btConsultar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "frmMayorGeneral";
			this.Text = "Mayor General";
			this.Load += new System.EventHandler(this.frmMayorGeneral_Load);
			this.TextChanged += new System.EventHandler(this.frmMayorGeneral_TextChanged);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalDebe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btConsultar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFinal.Focus();
				return;
			}

			if (this.txtInicial.Text.ToString().Length == 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtInicial.Focus();
				return;
			}
		
			bool bRango = false;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtInicial.Text.ToString();
				string sFinal = this.txtFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (int.Parse(sInicial) > int.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtInicial.Focus();
					return;
				}

				bRango = true;
			}
			
			bool bCta = false;

			if (this.txtCuenta.Text.ToString().Length > 0) bCta = true;

			this.Cursor = Cursors.WaitCursor;

			this.ultraGrid1.DataSource = ultraDataSource1;

			string sSQL = string.Format("Exec GeneraMayorGeneral '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtInicial.Text.ToString(), this.txtFinal.Text.ToString(), bRango, this.txtCuenta.Text.ToString(), bCta);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQL);

			this.Cursor = Cursors.Default;

			if (this.ultraGrid1.Rows.Count > 0) this.btnImprimir.Enabled = true; else this.btnImprimir.Enabled = false;

//			this.txtTotalDebe.Value = 0;
//				
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{
//				if (dr.Cells["Debe"].Value != System.DBNull.Value) this.txtTotalDebe.Value = Math.Round((double)this.txtTotalDebe.Value + (double)dr.Cells["Debe"].Value, 2);	
//			}
//
//			this.txtTotalHaber.Value = 0;
//
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{
//				if (dr.Cells["Haber"].Value != System.DBNull.Value) this.txtTotalHaber.Value = Math.Round((double)this.txtTotalHaber.Value + (double)dr.Cells["Haber"].Value, 2);	
//			}
//
//			this.txtSaldoTotal.Value = 0;
//				
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{
//				if (dr.Cells["Saldo"].Value != System.DBNull.Value) this.txtSaldoTotal.Value = Math.Round((double)this.txtSaldoTotal.Value + (double)dr.Cells["Saldo"].Value, 2);	
//			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			//this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Codigo", false, false);
			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);			
			
			//this.ultraGrid1.DisplayLayout.Bands[0].Expandable = true;						
		}

		private void frmMayorGeneral_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Asiento Asientos = new Asiento(((int)this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value));
			Asientos.MdiParent = MdiParent;
			Asientos.Show();
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			this.btnAsiento_Click(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtInicial.Focus();
		}

		private void txtInicial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtFinal.Enabled)this.txtFinal.Focus();
				else this.txtCuenta.Focus();
			}
		}

		private void txtFinal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCuenta.Enabled)this.txtCuenta.Focus();
				else this.btConsultar_Click(sender, e);
			}
		}

		private void txtCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btConsultar_Click(sender, e);
		}

		private void frmMayorGeneral_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtInicial_TextChanged(object sender, System.EventArgs e)
		{
			if (this.txtInicial.Text.Length > 0) 
			{
				this.txtFinal.Enabled = true;
				this.txtCuenta.Enabled = false;
			}
			else if (this.txtInicial.Text.Length == 0) 
			{
				this.txtFinal.Enabled = false;
				this.txtCuenta.Enabled = true;
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFinal.Focus();
				return;
			}

			if (this.txtInicial.Text.ToString().Length == 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtInicial.Focus();
				return;
			}
		
			bool bRango = false;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtInicial.Text.ToString();
				string sFinal = this.txtFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (int.Parse(sInicial) > int.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtInicial.Focus();
					return;
				}

				bRango = true;
			}
			
			bool bCta = false;

			if (this.txtCuenta.Text.ToString().Length > 0) bCta = true;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = dtDesde;//.ToString("yyyyMMdd"); 
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = dtHasta;//.ToString("yyyyMMdd"); 
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);

			ParameterField paramFieldInicial = new ParameterField ();
			ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
			paramFieldInicial.ParameterFieldName = "@Inicial";
			discreteValInicial.Value = this.txtInicial.Text.ToString(); 
			paramFieldInicial.CurrentValues.Add (discreteValInicial);
			paramFields.Add (paramFieldInicial);

			ParameterField paramFieldFinal = new ParameterField ();
			ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
			paramFieldFinal.ParameterFieldName = "@Final";
			discreteValFinal.Value = this.txtFinal.Text.ToString(); 
			paramFieldFinal.CurrentValues.Add (discreteValFinal);
			paramFields.Add (paramFieldFinal);

			ParameterField paramFieldRango = new ParameterField ();
			ParameterDiscreteValue discreteValRango = new ParameterDiscreteValue ();
			paramFieldRango.ParameterFieldName = "@Rango";
			discreteValRango.Value = bRango; 
			paramFieldRango.CurrentValues.Add (discreteValRango);
			paramFields.Add (paramFieldRango);

			ParameterField paramFieldCuenta = new ParameterField ();
			ParameterDiscreteValue discreteValCuenta = new ParameterDiscreteValue ();
			paramFieldCuenta.ParameterFieldName = "@Cuenta";
			discreteValCuenta.Value = this.txtCuenta.Text.ToString(); 
			paramFieldCuenta.CurrentValues.Add (discreteValCuenta);
			paramFields.Add (paramFieldCuenta);

			ParameterField paramFieldCta = new ParameterField ();
			ParameterDiscreteValue discreteValCta = new ParameterDiscreteValue ();
			paramFieldCta.ParameterFieldName = "@Cta";
			discreteValCta.Value = bCta; 
			paramFieldCta.CurrentValues.Add (discreteValCta);
			paramFields.Add (paramFieldCta);

			Reporte miRep = new Reporte("MayorGeneral.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Anticipo");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
		}
	}
}

