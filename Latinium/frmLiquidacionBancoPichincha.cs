using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmLiquidacionBancoPichincha.
	/// </summary>
	public class frmLiquidacionBancoPichincha : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsiento;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVerAsiento;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnContabilizar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLiquidacionBancoPichincha()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLiquidacionBancoPichincha));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num_Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion_GAF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Locales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Vendidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVendido");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contribucion_Solca");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalVendido", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalVendido", 12, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 13, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Contribucion_Solca", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Contribucion_Solca", 14, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 15, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num_Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion_GAF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Vendidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contribucion_Solca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtidAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.btnVerAsiento = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(696, 66);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(97, 16);
			this.label2.TabIndex = 641;
			this.label2.Text = "Fecha De Compra:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Enabled = false;
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(808, 64);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 640;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance1;
			this.uGridLotes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 75;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 59;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Factura N.";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 101;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Operación";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Cedula";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 266;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Liquido";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Venta";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "C Vendidas";
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance3;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Capital";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance4;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance5;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Impuesto";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance6;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Acreditado";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
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
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
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
			ultraGridBand1.SummaryFooterCaption = "";
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
			appearance16.FontData.SizeInPoints = 8F;
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
			this.uGridLotes.Location = new System.Drawing.Point(7, 104);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(1018, 272);
			this.uGridLotes.TabIndex = 638;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// txtLote
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance19;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(104, 8);
			this.txtLote.MinValue = -1;
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(88, 22);
			this.txtLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLote.TabIndex = 633;
			this.txtLote.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLote_KeyDown);
			this.txtLote.Click += new System.EventHandler(this.txtLote_Click);
			this.txtLote.ValueChanged += new System.EventHandler(this.txtLote_ValueChanged);
			this.txtLote.Enter += new System.EventHandler(this.txtLote_Enter);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(73, 16);
			this.label3.TabIndex = 637;
			this.label3.Text = "Lote Número:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 92);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1032, 8);
			this.groupBox1.TabIndex = 644;
			this.groupBox1.TabStop = false;
			// 
			// txtidAsiento
			// 
			this.txtidAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidAsiento.Appearance = appearance20;
			this.txtidAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtidAsiento.Enabled = false;
			this.txtidAsiento.FormatString = "";
			this.txtidAsiento.Location = new System.Drawing.Point(608, 8);
			this.txtidAsiento.Name = "txtidAsiento";
			this.txtidAsiento.PromptChar = ' ';
			this.txtidAsiento.Size = new System.Drawing.Size(8, 22);
			this.txtidAsiento.TabIndex = 645;
			this.txtidAsiento.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(8, 64);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 647;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(296, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 646;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(744, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 758;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(936, 63);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 757;
			this.btnContabilizar.Text = "Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// btnVerAsiento
			// 
			this.btnVerAsiento.CausesValidation = false;
			this.btnVerAsiento.Enabled = false;
			this.btnVerAsiento.Image = ((System.Drawing.Image)(resources.GetObject("btnVerAsiento.Image")));
			this.btnVerAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerAsiento.Location = new System.Drawing.Point(840, 8);
			this.btnVerAsiento.Name = "btnVerAsiento";
			this.btnVerAsiento.Size = new System.Drawing.Size(88, 23);
			this.btnVerAsiento.TabIndex = 756;
			this.btnVerAsiento.Text = "Ver Asiento";
			this.btnVerAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVerAsiento.Click += new System.EventHandler(this.btnVerAsiento_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(936, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 755;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(208, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 759;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 48);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1032, 8);
			this.groupBox2.TabIndex = 760;
			this.groupBox2.TabStop = false;
			// 
			// frmLiquidacionBancoPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1032, 382);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.btnVerAsiento);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtidAsiento);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.uGridLotes);
			this.Controls.Add(this.txtLote);
			this.Name = "frmLiquidacionBancoPichincha";
			this.Text = " ";
			this.Load += new System.EventHandler(this.frmLiquidacionBancoPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		public bool ActualizaNumeroOperacion(string stArchivo) 
		{ 			 
			if(File.Exists(stArchivo))
			{
				OleDbConnection oConn = new OleDbConnection(); 
				OleDbCommand oCmd = new OleDbCommand(); 
				OleDbDataAdapter oDa = new OleDbDataAdapter(); 
				DataSet oDs = new DataSet();
				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + stArchivo + "; Extended Properties= Excel 8.0;"; 
				oConn.Open(); 
				oCmd.CommandText = "SELECT * FROM [Lote$]"; 
				oCmd.Connection = oConn; 
				oDa.SelectCommand = oCmd; 
				oDa.Fill(oDs, "Lote");
 
				foreach (DataRow row in oDs.Tables[0].Rows)
				{					
					string sNumeroDeOperacion = "";
					string sCedula = "";
					decimal dCapital = 0.00m;
										
					if (row["OPERACION"] != System.DBNull.Value) sNumeroDeOperacion = row["OPERACION"].ToString();
					if (row["IDENTIFICACION"] != System.DBNull.Value) sCedula = row["IDENTIFICACION"].ToString();
					if (row["MONTO"] != System.DBNull.Value) dCapital = decimal.Parse(row["MONTO"].ToString());
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{
						if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
						
						if (sCedula.Trim() == dr.Cells["Ruc"].Value.ToString().Trim())
						{
							if (dCapital == Convert.ToDecimal(dr.Cells["CapitalVendido"].Value))
							{
								dr.Cells["Operacion_GAF"].Value = sNumeroDeOperacion.Trim();

								break;
							}
						}
					}
				}

				return true; 
			}
			else
			{
				return false; 
			}
		}
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		public bool bCarga = false;

		private void frmLiquidacionBancoPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320201");

			if (!Funcion.Permiso("961", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso LIQUIDACION CARTERA SOLIDARIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
 //And ISNULL(idAsientoLoteSolidario, 0) > 0
			
			bCarga = true;
			this.txtLote.Value = Funcion.iEscalarSQL(cdsSeteoF, "Cre_RetornaUltimoLote 4");
			this.txtLote.MaxValue = (int)this.txtLote.Value;
			bCarga = false;

				//Funcion.iEscalarSQL(cdsSeteoF, "Select MAX(ISNULL(NumeroLote, 0)) + 1 From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And idFormaPago = 9 And idEntidadFinanciera = 4 And Confirmado = 1");
			//if ((int)this.txtLote.Value == 0) this.txtLote.Value = (int)this.txtLote.MaxValue;

//			this.txtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select MAX(ISNULL(FechaVenta, 0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And idFormaPago = 9 And idEntidadFinanciera = 4 And Confirmado = 1");

			FuncionesProcedimientos.RetornaFechaServidor(this.txtFecha, cdsSeteoF, false);
		}

		private void txtLote_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridLotes);
		}

		private void btnVerAsiento_Click(object sender, System.EventArgs e)
		{			
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("No hay Documentos asignados en el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
				this.txtLote.Focus();
				return;
			}

			if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show(string.Format("No se ha Contabilizado el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}
			
			Cursor = Cursors.WaitCursor;
			frmAsientos Asiento = new frmAsientos(((int)this.txtidAsiento.Value));
			Asiento.MdiParent = MdiParent;
			Asiento.Show();
			Cursor = Cursors.Default;			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("No hay Documentos asignados en el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}

			if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show(string.Format("No se ha Contabilizado el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}
			
			#region Impresion
			string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();
				
			Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Asiento");			
			miRepor.Show();	
			#endregion Impresion			
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{			
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Documentos con este Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}

//			int iUltimoLoteLiq = Funcion.iEscalarSQL(cdsSeteoF, "Select MAX(ISNULL(NumeroLote, 0)) + 1 From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And idFormaPago = 9 And idEntidadFinanciera = 4 And CapturaLote = 1 And ISNULL(idAsientoLoteSolidario, 0) > 0");
//
//			if ((int)this.txtLote.Value > iUltimoLoteLiq)
//			{
//				MessageBox.Show(string.Format("El Lote {0} esta pendiente de liquidar", iUltimoLoteLiq), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
//				this.txtLote.Focus();
//				return;
//			}

			if (!Validacion.vbFechaEnDocumentos(this.txtFecha, "Ingrese la fecha del documento", true, "Bancos", cdsSeteoF)) return;
			
			DateTime dtUltLiq = Funcion.dtEscalarSQL(cdsSeteoF, "Select MAX(ISNULL(FechaVenta, 0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And idFormaPago = 9 And idEntidadFinanciera = 4 And CapturaLote = 1");

			if (Convert.ToDateTime(this.txtFecha.Value) < dtUltLiq)
			{				
				MessageBox.Show(string.Format("La Fecha de Liquidación no puede ser anterior a '{0}'", dtUltLiq.ToString("yyyy/MM/dd")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFecha.Value = dtUltLiq;
				this.txtFecha.Focus();			
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de Contabilizar el Lote Numero {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{		
				this.openFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName.Length > 0) 
				{
					this.Cursor = Cursors.WaitCursor;

					Funcion.TiempoInicia();
					
					this.ActualizaNumeroOperacion(this.openFileDialog1.FileName); 

					decimal dFondo = 0.00m;
					decimal dImpuesto = 0.00m;
					decimal dAcreditado = 0.00m;
					int iCont = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{
						dFondo = dFondo + Convert.ToDecimal(dr.Cells["Comision"].Value);
						dImpuesto = dImpuesto + Convert.ToDecimal(dr.Cells["Contribucion_Solca"].Value);
						dAcreditado = dAcreditado + Convert.ToDecimal(dr.Cells["Valor"].Value);
					}
										
					string sSQL = string.Format("Exec AsientoLiquidacionCarteraBancoPichincha {0}, {1}, '{2}', {3}, {4}, {5}", 
						(int)this.txtidAsiento.Value, (int)this.txtLote.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), dFondo, dImpuesto, dAcreditado);
					this.txtidAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					
					string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenLiquidacionCartera {0}, {1}, '{2}', 4", 0, (int)this.txtLote.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCO);
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set idAsientoLoteSolidario = {0}, FechaVenta = '{1}', NumeroDeOperacion = '{2}', Bloquear = 1, Confirmado = 1, ValorAcreditado = {4}, Contribucion_Solca = {5}, Comision = {6} Where idCompra = {3}", 
							(int)this.txtidAsiento.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), dr.Cells["Operacion_GAF"].Value.ToString(), (int)dr.Cells["idCompra"].Value,
							Convert.ToDecimal(dr.Cells["Valor"].Value), Convert.ToDecimal(dr.Cells["Contribucion_Solca"].Value), Convert.ToDecimal(dr.Cells["Comision"].Value))); 
											
						iCont++;
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_LiquidaTablasCarteraBP {0}, '{1}'", (int)this.txtLote.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd")));

					MessageBox.Show(string.Format("Se contabilizó {1} documentos del Lote {0} " + Funcion.Tiempo("EN").ToString(), (int)this.txtLote.Value, iCont.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					#region Impresion
					string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();				
					Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("Asiento");			
					miRepor.Show();	
					#endregion Impresion

					this.btnVer_Click(sender, e);

					this.Cursor = Cursors.Default;
				}			
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}

			this.lblEstado.Text = "";
			this.lblContador.Text = "";

			this.txtFecha.Enabled = false;
			this.btnContabilizar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnVerAsiento.Enabled = false;

			string sSQL = string.Format("Exec Cre_LiquidacionCarteraBancoPichincha {0}", (int)this.txtLote.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridLotes);
						
			this.txtidAsiento.Value = 
				Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idAsiento Int = 0 If Exists(Select ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idOrigenAsiento = 49 And idTipoAsiento = 1 And Borrar = 0) Select @idAsiento = ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idTipoAsiento = 1 And Borrar = 0 Select @idAsiento", this.txtLote.Value.ToString()));

			if ((int)this.txtidAsiento.Value > 0)
			{				
				this.lblEstado.Text = "CONTABILIZADO";				
				this.btnImprimir.Enabled = true;
				this.btnVerAsiento.Enabled = true;
			}
			else 
			{
				this.lblEstado.Text = "CONTABILIZACIÓN PENDIENTE";				
				this.txtFecha.Enabled = true;
				this.btnContabilizar.Enabled = true;				
			}		

			this.lblContador.Text = this.uGridLotes.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void txtLote_Enter(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll();
		}

		private void txtLote_Click(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll();
		}

		private void txtLote_ValueChanged(object sender, System.EventArgs e)
		{
			if (bCarga) this.btnVer_Click(sender, e);
		}
	}
}

