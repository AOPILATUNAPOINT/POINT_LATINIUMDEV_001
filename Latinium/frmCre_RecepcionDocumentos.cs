using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCre_RecepcionDocumentos.
	/// </summary>
	public class frmCre_RecepcionDocumentos : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCreditos;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_RecepcionDocumentos()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		
		int i = 0;
		string sRUC = "";
		DateTime dtFecha = DateTime.Today;

		public frmCre_RecepcionDocumentos(int I, string RUC, DateTime Fecha)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			i = I;
			sRUC = RUC;
			dtFecha = Fecha;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_RecepcionDocumentos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagare");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Amortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GastosDeCobranza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LugarDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresosGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ReservaDeDominio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CheckList");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNotas", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 7, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagare");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Amortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GastosDeCobranza");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LugarDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IngresosGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ReservaDeDominio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CheckList");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seguro");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridCreditos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 370;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 16);
			this.label5.TabIndex = 369;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 0;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// uGridCreditos
			// 
			this.uGridCreditos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCreditos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCreditos.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridCreditos.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 96;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "F. Factura";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Nota del Sistema";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Recepcion";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 56;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 50;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.Header.Caption = "Cobranza";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 60;
			ultraGridColumn18.Header.Caption = "Lugar Pago";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 70;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 50;
			ultraGridColumn20.Header.Caption = "I/G";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 50;
			ultraGridColumn21.Header.Caption = "Dominio";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 60;
			ultraGridColumn22.Header.Caption = "Check List";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 70;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn24.Width = 400;
			ultraGridColumn25.Header.VisiblePosition = 22;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			ultraGridColumn26.CellButtonAppearance = appearance5;
			ultraGridColumn26.Header.Caption = "Notas";
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 50;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridCreditos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCreditos.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.uGridCreditos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uGridCreditos.DisplayLayout.Override.CardAreaAppearance = appearance8;
			this.uGridCreditos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCreditos.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridCreditos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridCreditos.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridCreditos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCreditos.Location = new System.Drawing.Point(8, 88);
			this.uGridCreditos.Name = "uGridCreditos";
			this.uGridCreditos.Size = new System.Drawing.Size(1248, 240);
			this.uGridCreditos.TabIndex = 371;
			this.uGridCreditos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCreditos_CellChange);
			this.uGridCreditos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCreditos_ClickCellButton);
			this.uGridCreditos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCreditos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn12.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn18.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn19.DefaultValue = false;
			ultraDataColumn20.DataType = typeof(bool);
			ultraDataColumn20.DefaultValue = false;
			ultraDataColumn21.DataType = typeof(bool);
			ultraDataColumn21.DefaultValue = false;
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn22.DefaultValue = false;
			ultraDataColumn25.DataType = typeof(bool);
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 373;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance12;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 40);
			this.txtBuscar.MaxLength = 15;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(288, 22);
			this.txtBuscar.TabIndex = 2;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnGenerar
			// 
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(368, 40);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.TabIndex = 3;
			this.btnGenerar.Text = "&Ver";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1264, 8);
			this.groupBox1.TabIndex = 375;
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.CausesValidation = false;
			this.btnExcel.Enabled = false;
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1179, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 387;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(480, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 388;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optEstados
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance13;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance14;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 2;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Completado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(480, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(240, 24);
			this.optEstados.TabIndex = 638;
			this.optEstados.Text = "Pendiente";
			// 
			// frmCre_RecepcionDocumentos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1264, 334);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridCreditos);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnExcel);
			this.KeyPreview = true;
			this.Name = "frmCre_RecepcionDocumentos";
			this.ShowInTaskbar = false;
			this.Text = "Recepción de Documentos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_RecepcionDocumentos_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_RecepcionDocumentos_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQL = string.Format("Exec Cre_RecepcionDeDocumentos '{0}', '{1}', '{2}', {3}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(), Convert.ToInt32(this.optEstados.Value));
			this.uGridCreditos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridCreditos.Rows.Count + " Registros Encontrados";
			#region Acceso Editar
			if (!miAcceso.Editar)
			{
	
//			this.uGridCreditos.DisplayLayout.Bands[0].Columns["btnNotas"].CellActivation  = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				this.uGridCreditos.DisplayLayout.Bands[0].Columns["btnNotas"].Hidden = true;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Notas"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Usuario"].CellActivation  = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["CheckList"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["ReservaDeDominio"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["IngresosGastos"].CellActivation  = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Factura"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["LugarDePago"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["GastosDeCobranza"].CellActivation  = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Amortizacion"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Pagare"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Cedula"].CellActivation  = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

				this.uGridCreditos.DisplayLayout.Bands[0].Columns["Solicitud"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

			}
			#endregion Acceso Editar
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCre_RecepcionDocumentos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0822");

			if (!Funcion.Permiso("629", cdsSeteoF))
			{				
				MessageBox.Show("No tiene acceso a Recepción de Documentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BExportar)this.btnExcel.Enabled = true;

			DateTime dtFecha = DateTime.Parse("25/02/2016");
			this.dtDesde.CalendarInfo.MinDate = dtFecha.Date;

			this.dtHasta.Value = DateTime.Today;

			if (i == 1)
			{
				this.dtDesde.Value = dtFecha;
				this.txtBuscar.Text = sRUC;
				this.btnGenerar_Click(sender, e);
			}
			else
			{
				this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			}	
			
		
		}

		private void uGridCreditos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Solicitud
			if (e.Cell.Column.ToString() == "Solicitud")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Solicitud', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["Solicitud"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));				
			}
			#endregion Solicitud

			#region Cedula
			if (e.Cell.Column.ToString() == "Cedula")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Cedula', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["Cedula"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));				
			}
			#endregion Cedula

			#region Pagare
			if (e.Cell.Column.ToString() == "Pagare")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Pagare', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["Pagare"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion Pagare

			#region Amortizacion
			if (e.Cell.Column.ToString() == "Amortizacion")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Amortizacion', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["Amortizacion"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion Amortizacion

			#region GastosDeCobranza
			if (e.Cell.Column.ToString() == "GastosDeCobranza")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'GastosDeCobranza', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["GastosDeCobranza"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion GastosDeCobranza

			#region LugarDePago
			if (e.Cell.Column.ToString() == "LugarDePago")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'LugarDePago', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["LugarDePago"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion LugarDePago

			#region Factura
			if (e.Cell.Column.ToString() == "Factura")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Factura', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["Factura"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion Factura

			#region IngresosGastos
			if (e.Cell.Column.ToString() == "IngresosGastos")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'IngresosGastos', {0}, {1}", Convert.ToBoolean(e.Cell.Row.Cells["IngresosGastos"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion IngresosGastos

			#region ReservaDeDominio
			if (e.Cell.Column.ToString() == "ReservaDeDominio")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'ReservaDeDominio', {0}, {1}",  Convert.ToBoolean(e.Cell.Row.Cells["ReservaDeDominio"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion ReservaDeDominio

			#region Check List
			if (e.Cell.Column.ToString() == "CheckList")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'CheckList', {0}, {1}",  Convert.ToBoolean(e.Cell.Row.Cells["CheckList"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion Check List

			#region Seguro
			if (e.Cell.Column.ToString() == "Seguro")
			{
				this.uGridCreditos.UpdateData();
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaRecepcionDocumentos 'Seguro', {0}, {1}",  Convert.ToBoolean(e.Cell.Row.Cells["Seguro"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCompra"].Value)));
			}
			#endregion Seguro
		}

		private void uGridCreditos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnNotas")
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(0))
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						try
						{
							int iCont = 0;

							for (int j = 14; j <= 23; j++)
							{
								if (Convert.ToBoolean(e.Cell.Row.Cells[j].Value)) iCont++;
							}
														
							string sSQL = string.Format("Exec Cre_ActualizaEstadoCreditoRD {0}, '{1}', {2}", (int)e.Cell.Row.Cells["idCompra"].Value, Nota.txtNotas.Text.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error del Sistema : {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void uGridCreditos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGenerar_Click(sender, e);
		}

		private void frmCre_RecepcionDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridCreditos.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.uGridCreditos, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}
	}
}

