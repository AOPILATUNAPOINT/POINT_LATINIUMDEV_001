using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Summary description for frmReporteCuentasPorPagar.
	/// </summary>
	public class frmReporteCuentasPorPagar : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.RadioGroup rdgEstado;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.RadioGroup rdgGrupo;
		private System.Windows.Forms.Button btnActualizarSaldos;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnImpresion;
		private System.Windows.Forms.Button btnAbrir;
		private DevExpress.XtraEditors.RadioGroup rgOrigen;
		private DevExpress.XtraEditors.RadioGroup rgDocumentos;
		private System.Windows.Forms.Button btnIrA;
		private System.Windows.Forms.ContextMenu mnuIrA;
		private System.Windows.Forms.MenuItem mnuIrADocumento;
		private System.Windows.Forms.MenuItem mnuIrAEstadoDeCuenta;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteCuentasPorPagar()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCuentasPorPagar));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenM");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIVA0");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIVA12");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagado");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAnticipo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 19, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 21, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pagado", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pagado", 22, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoAnticipo", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoAnticipo", 23, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIVA0", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIVA0", 16, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIVA12", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIVA12", 17, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 18, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OrigenM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIVA0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIVA12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.rdgEstado = new DevExpress.XtraEditors.RadioGroup();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.rdgGrupo = new DevExpress.XtraEditors.RadioGroup();
			this.btnActualizarSaldos = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnImpresion = new System.Windows.Forms.Button();
			this.btnAbrir = new System.Windows.Forms.Button();
			this.rgOrigen = new DevExpress.XtraEditors.RadioGroup();
			this.rgDocumentos = new DevExpress.XtraEditors.RadioGroup();
			this.btnIrA = new System.Windows.Forms.Button();
			this.mnuIrA = new System.Windows.Forms.ContextMenu();
			this.mnuIrADocumento = new System.Windows.Forms.MenuItem();
			this.mnuIrAEstadoDeCuenta = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rgOrigen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rgDocumentos.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(14, 16);
			this.label1.TabIndex = 638;
			this.label1.Text = "Al";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 632;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(64, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 2;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtHasta_KeyPress);
			// 
			// txtBuscar
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance2;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(176, 22);
			this.txtBuscar.TabIndex = 3;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 51;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Header.Caption = "Origen";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 40;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 47;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 53;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 40;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Vence";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "R.U.C.";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Proveedor";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 250;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 102;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 66;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.Caption = "Forma de Pago";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 99;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance6;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "BI IVA 0";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance7;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "BI IVA 12";
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance8;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance9;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 60;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance10;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance11;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance12;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Saldo";
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Width = 80;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance13;
			ultraGridColumn25.Header.Caption = "D Mora";
			ultraGridColumn25.Header.VisiblePosition = 16;
			ultraGridColumn25.Width = 50;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance14;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 100;
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
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance17;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.Hidden = true;
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance18;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance19;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance20;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance21;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 8F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 128);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1088, 296);
			this.uGridDocumentos.TabIndex = 631;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(int);
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
																																 ultraDataColumn25,
																																 ultraDataColumn26});
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
			// rdgEstado
			// 
			this.rdgEstado.EditValue = 1;
			this.rdgEstado.Location = new System.Drawing.Point(360, 8);
			this.rdgEstado.Name = "rdgEstado";
			// 
			// rdgEstado.Properties
			// 
			this.rdgEstado.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgEstado.Properties.Appearance.Options.UseBackColor = true;
			this.rdgEstado.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.rdgEstado.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Pendiente"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Cancelado")});
			this.rdgEstado.Size = new System.Drawing.Size(88, 96);
			this.rdgEstado.TabIndex = 0;
			this.rdgEstado.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgEstado_KeyPress);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(856, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(31, 16);
			this.label5.TabIndex = 647;
			this.label5.Text = "Saldo";
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(856, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(29, 16);
			this.label4.TabIndex = 646;
			this.label4.Text = "Total";
			// 
			// txtSaldo
			// 
			this.txtSaldo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance29;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(904, 7);
			this.txtSaldo.MinValue = -2147483647;
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(88, 22);
			this.txtSaldo.TabIndex = 645;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance30;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(904, 40);
			this.txtTotal.MinValue = -2147483647;
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 22);
			this.txtTotal.TabIndex = 644;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 648;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1104, 8);
			this.groupBox1.TabIndex = 649;
			this.groupBox1.TabStop = false;
			// 
			// rdgGrupo
			// 
			this.rdgGrupo.EditValue = 0;
			this.rdgGrupo.Location = new System.Drawing.Point(456, 8);
			this.rdgGrupo.Name = "rdgGrupo";
			// 
			// rdgGrupo.Properties
			// 
			this.rdgGrupo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgGrupo.Properties.Appearance.Options.UseBackColor = true;
			this.rdgGrupo.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.rdgGrupo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(-1, "Locales - Servicios"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Locales"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(5, "Servicios"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Asia"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Internacionales")});
			this.rdgGrupo.Size = new System.Drawing.Size(120, 96);
			this.rdgGrupo.TabIndex = 653;
			// 
			// btnActualizarSaldos
			// 
			this.btnActualizarSaldos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizarSaldos.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizarSaldos.Image")));
			this.btnActualizarSaldos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizarSaldos.Location = new System.Drawing.Point(1008, 80);
			this.btnActualizarSaldos.Name = "btnActualizarSaldos";
			this.btnActualizarSaldos.Size = new System.Drawing.Size(88, 24);
			this.btnActualizarSaldos.TabIndex = 656;
			this.btnActualizarSaldos.Text = "Saldos";
			this.btnActualizarSaldos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizarSaldos.Click += new System.EventHandler(this.btnActualizarSaldos_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1008, 56);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 658;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(256, 39);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(75, 24);
			this.btnBuscar.TabIndex = 657;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnImpresion
			// 
			this.btnImpresion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImpresion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImpresion.CausesValidation = false;
			this.btnImpresion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImpresion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImpresion.Image = ((System.Drawing.Image)(resources.GetObject("btnImpresion.Image")));
			this.btnImpresion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImpresion.Location = new System.Drawing.Point(1008, 32);
			this.btnImpresion.Name = "btnImpresion";
			this.btnImpresion.Size = new System.Drawing.Size(88, 24);
			this.btnImpresion.TabIndex = 731;
			this.btnImpresion.Text = "Imprimir";
			this.btnImpresion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImpresion.Click += new System.EventHandler(this.btnImpresion_Click);
			// 
			// btnAbrir
			// 
			this.btnAbrir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAbrir.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrir.Image")));
			this.btnAbrir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAbrir.Location = new System.Drawing.Point(904, 72);
			this.btnAbrir.Name = "btnAbrir";
			this.btnAbrir.Size = new System.Drawing.Size(88, 24);
			this.btnAbrir.TabIndex = 732;
			this.btnAbrir.Text = "Ver Factura";
			this.btnAbrir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAbrir.Click += new System.EventHandler(this.btnAbrir_Click);
			// 
			// rgOrigen
			// 
			this.rgOrigen.EditValue = -1;
			this.rgOrigen.Location = new System.Drawing.Point(584, 8);
			this.rgOrigen.Name = "rgOrigen";
			// 
			// rgOrigen.Properties
			// 
			this.rgOrigen.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rgOrigen.Properties.Appearance.Options.UseBackColor = true;
			this.rgOrigen.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
			this.rgOrigen.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(-1, "Todos"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Facturación"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Caja Chica"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Tarjeta de Crédito")});
			this.rgOrigen.Size = new System.Drawing.Size(120, 96);
			this.rgOrigen.TabIndex = 733;
			// 
			// rgDocumentos
			// 
			this.rgDocumentos.EditValue = 0;
			this.rgDocumentos.Location = new System.Drawing.Point(712, 8);
			this.rgDocumentos.Name = "rgDocumentos";
			// 
			// rgDocumentos.Properties
			// 
			this.rgDocumentos.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rgDocumentos.Properties.Appearance.Options.UseBackColor = true;
			this.rgDocumentos.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
			this.rgDocumentos.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Facturas"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(6, "Notas De Crédito"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(28, "Notas de Debito"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(25, "Anticipos")});
			this.rgDocumentos.Size = new System.Drawing.Size(112, 96);
			this.rgDocumentos.TabIndex = 734;
			// 
			// btnIrA
			// 
			this.btnIrA.Image = ((System.Drawing.Image)(resources.GetObject("btnIrA.Image")));
			this.btnIrA.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIrA.Location = new System.Drawing.Point(1008, 8);
			this.btnIrA.Name = "btnIrA";
			this.btnIrA.Size = new System.Drawing.Size(88, 24);
			this.btnIrA.TabIndex = 735;
			this.btnIrA.Text = "Ir a";
			this.btnIrA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIrA.Click += new System.EventHandler(this.btnIrA_Click);
			// 
			// mnuIrA
			// 
			this.mnuIrA.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuIrADocumento,
																																					 this.mnuIrAEstadoDeCuenta});
			// 
			// mnuIrADocumento
			// 
			this.mnuIrADocumento.Index = 0;
			this.mnuIrADocumento.Text = "Documento";
			// 
			// mnuIrAEstadoDeCuenta
			// 
			this.mnuIrAEstadoDeCuenta.Index = 1;
			this.mnuIrAEstadoDeCuenta.Text = "Estado De Cuenta";
			this.mnuIrAEstadoDeCuenta.Click += new System.EventHandler(this.mnuIrAEstadoDeCuenta_Click);
			// 
			// frmReporteCuentasPorPagar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1104, 430);
			this.Controls.Add(this.btnIrA);
			this.Controls.Add(this.rgDocumentos);
			this.Controls.Add(this.rgOrigen);
			this.Controls.Add(this.btnAbrir);
			this.Controls.Add(this.btnImpresion);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnActualizarSaldos);
			this.Controls.Add(this.rdgGrupo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtSaldo);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.rdgEstado);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridDocumentos);
			this.KeyPreview = true;
			this.Name = "frmReporteCuentasPorPagar";
			this.Text = "Resumen de Cuentas Por Pagar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteCuentasPorPagar_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteCuentasPorPagar_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rgOrigen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rgDocumentos.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBuscar.Focus();
		}

		private void rdgEstado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBuscar.Focus();
		}

		private void frmReporteCuentasPorPagar_Load(object sender, System.EventArgs e)
		{
			//this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtHasta.Value = DateTime.Today;
		}

		private void frmReporteCuentasPorPagar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}
		
		private void btnActualizarSaldos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No hay documentos para actualizar saldos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();			
			}

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Actualizar Saldos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
			{	
				this.Cursor = Cursors.WaitCursor;
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, {1}, 0", (int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoFactura"].Value));
				}
				
				this.Cursor = Cursors.Default;	

				MessageBox.Show("Saldos Actualizados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnBuscar_Click(sender, e);
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ReporteCuentaPorPagar '{0}', '{1}', {2}, {3}, {4}, {5}", 
				this.txtBuscar.Text.ToString(), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)this.rdgEstado.EditValue, (int)this.rdgGrupo.EditValue, (int)this.rgOrigen.EditValue, (int)this.rgDocumentos.EditValue);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);

			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " DOCUMENTOS";

			this.txtTotal.Value = 0.00m;
			this.txtSaldo.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				this.txtTotal.Value = Convert.ToDecimal(this.txtTotal.Value) + Convert.ToDecimal(dr.Cells["Total"].Value); 
				this.txtSaldo.Value = Convert.ToDecimal(this.txtSaldo.Value) + Convert.ToDecimal(dr.Cells["SaldoAnticipo"].Value); 
			}
			
			//			this.uGridDocumentos.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
			//			this.uGridDocumentos.Rows.ExpandAll(true);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void btnImpresion_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfCapital = new ParameterField ();
			ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
			pfCapital.ParameterFieldName = "@Filtro";
			ValCapital.Value = this.txtBuscar.Text.ToString();
			pfCapital.CurrentValues.Add (ValCapital);
			paramFields.Add (pfCapital);
			
			ParameterField pfFecha = new ParameterField ();
			ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
			pfFecha.ParameterFieldName = "@Hasta";
			ValFecha.Value = dtHasta;
			pfFecha.CurrentValues.Add (ValFecha);
			paramFields.Add (pfFecha);

			ParameterField pfCuotas = new ParameterField ();
			ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
			pfCuotas.ParameterFieldName = "@Estado";
			ValCuotas.Value = (int)this.rdgEstado.EditValue;
			pfCuotas.CurrentValues.Add (ValCuotas);
			paramFields.Add (pfCuotas);
				
			ParameterField pfGrupo = new ParameterField ();
			ParameterDiscreteValue ValGrupo = new ParameterDiscreteValue ();
			pfGrupo.ParameterFieldName = "@idGrupoCliente";
			ValGrupo.Value = (int)this.rdgGrupo.EditValue;
			pfGrupo.CurrentValues.Add (ValGrupo);
			paramFields.Add (pfGrupo);
						
			Reporte miRepor = new Reporte("EstadoCuentasPorPagar.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	
		}

		private void btnAbrir_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros en la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridDocumentos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				frmCompra Compra = new frmCompra((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, true);
				Compra.MdiParent = MdiParent;
				Compra.Show();
			}
		}

		private void mnuIrAEstadoDeCuenta_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmCuentasPorPagar CXP = new frmCuentasPorPagar(this.uGridDocumentos.ActiveRow.Cells["Ruc"].Value.ToString());
			CXP.MdiParent = MdiParent;
			CXP.Show();
		}

		private void btnIrA_Click(object sender, System.EventArgs e)
		{
			this.mnuIrA.Show(this.btnIrA, new Point(1, 20));
		}
	}
}

