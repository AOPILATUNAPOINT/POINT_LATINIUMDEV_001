

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmComparaCarteraPichincha.
	/// </summary>
	public class frmComparaCarteraPichincha : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCartera;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnAcciones;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnICVencida;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmComparaCarteraPichincha()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencidas");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagadas");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Por_Vencer");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No_Devenga");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCapital");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagos_Por_Enviar");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Envio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoPichincha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NC_RM");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Pic");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencido_Pic");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Por_Vencer_Pic");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoDevenga_Pic");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Canceladas_Pic");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencidas_Pic");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCapitalPP_Pic");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCapitalPV_Pic");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 7, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido", 8, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Por_Vencer", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Por_Vencer", 9, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "No_Devenga", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "No_Devenga", 10, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoCapital", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoCapital", 11, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pagos_Por_Enviar", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pagos_Por_Enviar", 12, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vencido_Pic", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vencido_Pic", 21, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Por_Vencer_Pic", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Por_Vencer_Pic", 22, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NoDevenga_Pic", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NoDevenga_Pic", 23, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoCapitalPP_Pic", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoCapitalPP_Pic", 26, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoCapitalPV_Pic", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoCapitalPV_Pic", 27, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Por_Vencer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No_Devenga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagos_Por_Enviar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Envio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoPichincha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NC_RM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Por_Vencer_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NoDevenga_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Canceladas_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencidas_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapitalPP_Pic");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapitalPV_Pic");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmComparaCarteraPichincha));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridCartera = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnAcciones = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.btnICVencida = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
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
			// uGridCartera
			// 
			this.uGridCartera.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCartera.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCartera.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCartera.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Operación";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Por Vencer";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "No Devenga";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance8;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Saldo Capital";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Pagos Por Enviar";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Ultimo Pago";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Ultimo Envio";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Format = "";
			ultraGridColumn17.Header.Caption = "Dias Mora";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Rango Pichincha";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 110;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellAppearance = appearance11;
			ultraGridColumn20.Format = "";
			ultraGridColumn20.Header.Caption = "NC RM";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance12;
			ultraGridColumn21.Format = "";
			ultraGridColumn21.Header.Caption = "Dias Mora Pichincha";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 110;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance13;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Vencido Pichincha";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 110;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance14;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Por Vencer Pichincha";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 110;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance15;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "No Devenga Pichincha";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 110;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance16;
			ultraGridColumn25.Header.Caption = "Canceladas Pichincha";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 110;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance17;
			ultraGridColumn26.Header.Caption = "Vencidas Pichincha";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 110;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance18;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "Saldo Capital PP Pichincha";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 110;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance19;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "Saldo Capital PV Pichincha";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 110;
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
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			appearance20.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance20;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance21;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance22;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance23;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance24;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance25;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance26;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance27;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance28;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance29;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance30;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance31;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridCartera.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCartera.DisplayLayout.Override.ActiveRowAppearance = appearance32;
			this.uGridCartera.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.uGridCartera.DisplayLayout.Override.CardAreaAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 8F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCartera.DisplayLayout.Override.HeaderAppearance = appearance34;
			this.uGridCartera.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridCartera.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridCartera.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCartera.Location = new System.Drawing.Point(8, 88);
			this.uGridCartera.Name = "uGridCartera";
			this.uGridCartera.Size = new System.Drawing.Size(1320, 272);
			this.uGridCartera.TabIndex = 200;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// txtBusqueda
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance37;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(104, 41);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 22);
			this.txtBusqueda.TabIndex = 273;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(16, 44);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(76, 15);
			this.ultraLabel1.TabIndex = 272;
			this.ultraLabel1.Text = "Cliente/Cedula";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1576, 8);
			this.groupBox1.TabIndex = 700;
			this.groupBox1.TabStop = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 17);
			this.label5.TabIndex = 753;
			this.label5.Text = "Fecha";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance38;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(104, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 752;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(408, 43);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1012;
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(312, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1013;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1248, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 1014;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnAcciones
			// 
			this.btnAcciones.Image = ((System.Drawing.Image)(resources.GetObject("btnAcciones.Image")));
			this.btnAcciones.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcciones.Location = new System.Drawing.Point(1248, 8);
			this.btnAcciones.Name = "btnAcciones";
			this.btnAcciones.Size = new System.Drawing.Size(75, 24);
			this.btnAcciones.TabIndex = 1015;
			this.btnAcciones.Text = "Ratio";
			this.btnAcciones.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcciones.Click += new System.EventHandler(this.btnAcciones_Click);
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1080, 40);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(152, 24);
			this.button1.TabIndex = 1016;
			this.button1.Text = "Importar Cartera Vigente";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnICVencida
			// 
			this.btnICVencida.Image = ((System.Drawing.Image)(resources.GetObject("btnICVencida.Image")));
			this.btnICVencida.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnICVencida.Location = new System.Drawing.Point(1080, 8);
			this.btnICVencida.Name = "btnICVencida";
			this.btnICVencida.Size = new System.Drawing.Size(152, 24);
			this.btnICVencida.TabIndex = 1017;
			this.btnICVencida.Text = "Importar Cartera Vencida";
			this.btnICVencida.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmComparaCarteraPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1336, 366);
			this.Controls.Add(this.btnICVencida);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnAcciones);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.uGridCartera);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmComparaCarteraPichincha";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Compara Cartera Pichincha";
			this.Load += new System.EventHandler(this.frmComparaCarteraPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

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
		private void frmComparaCarteraPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200606");

			if (!Funcion.Permiso("1315", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte Seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (!miAcceso.BExportar) this.btnExcel.Enabled = true;
			this.dtFecha.Value = DateTime.Today ;
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int idCarteraPichi = 0;
		
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			idCarteraPichi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select isnull (COUNT(*),0) from TEMPORALCARTERAPICHIS where CONVERT ( DATE, fecha) = '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));		

			if (idCarteraPichi == 0)
			{
				MessageBox.Show("Importe la Cartera Vigente Del Pichincha del dia Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.lblContador.Text = "";

			string sSQLPa = string.Format("Exec ReporteComparaCarteraPichincha '{0}', '{1}' ", 
			 Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString ());
		
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridCartera);
				this.uGridCartera.Visible = true;
			
			this.lblContador.Text = this.uGridCartera.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridCartera);
		}

		private void btnAcciones_Click(object sender, System.EventArgs e)
		{
			using (frmRatios Ratios = new frmRatios((DateTime)this.dtFecha.Value))
			{
				if (DialogResult.OK == Ratios.ShowDialog())
				{

				}				
			}	
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			int idCarteraPichi = 0;
		
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			idCarteraPichi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select isnull (COUNT(*),0) from TEMPORALCARTERAPICHIS where CONVERT ( DATE, fecha) = '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));		

			if (idCarteraPichi > 0)
			{
				MessageBox.Show("Existen datos importados a la fecha actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			#region Variables
			int iCont = 0;
			string	OPERACION	="";
			string	IDENTIFICACION	 ="";
			string	CLIENTE	="";
			decimal	MONTO_FINANCIADO	= 0;
			decimal	SALDO_CAPITAL_VENCIDO	= 0;
			decimal	SALDO_CAPITAL_X_VENCER =  0;	
			decimal	SALDO_CAPITAL_NODEV = 0;	
			decimal	INTERES_X_PAGAR	= 0;
			decimal	MORA_X_PAGAR	= 0;
			int	DIAS_MORA	= 0;
			decimal	NUMERO_CUOTAS_VENCIDAS	= 0;
			decimal	V_CUOTAS_VENCIDAS	= 0;
			int	CUOTAS_CANCELADAS	= 0;
			#endregion Variables

			#region Sube archivo
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
										
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)

					{					
						if (row["OPERACION"] != System.DBNull.Value)  OPERACION = row["OPERACION"].ToString();
						if (row["IDENTIFICACION"] != System.DBNull.Value) IDENTIFICACION = row["IDENTIFICACION"].ToString();
						if (row["CLIENTE"] != System.DBNull.Value) CLIENTE = row["CLIENTE"].ToString();
						if (row["MONTO_FINANCIADO"] != System.DBNull.Value) MONTO_FINANCIADO = decimal.Parse( row["MONTO_FINANCIADO"].ToString()); 
						if (row["SALDO_CAPITAL_VENCIDO"] != System.DBNull.Value) SALDO_CAPITAL_VENCIDO = decimal.Parse( row["SALDO_CAPITAL_VENCIDO"] .ToString());
						if (row["SALDO_CAPITAL_X_VENCER"] != System.DBNull.Value) SALDO_CAPITAL_X_VENCER = decimal.Parse( row["SALDO_CAPITAL_X_VENCER"].ToString());
						if (row["SALDO_CAPITAL_NODEV"] != System.DBNull.Value) SALDO_CAPITAL_NODEV = decimal.Parse( row["SALDO_CAPITAL_NODEV"] .ToString());
						if (row["INTERES_X_PAGAR"] != System.DBNull.Value) INTERES_X_PAGAR = decimal.Parse( row["INTERES_X_PAGAR"].ToString()); 
						if (row["MORA_X_PAGAR"] != System.DBNull.Value) MORA_X_PAGAR = decimal.Parse( row["MORA_X_PAGAR"].ToString()); 
						if (row["DIAS_MORA"] != System.DBNull.Value) DIAS_MORA = int.Parse( row["DIAS_MORA"].ToString()); 
						if (row["NUMERO_CUOTAS_VENCIDAS"] != System.DBNull.Value) NUMERO_CUOTAS_VENCIDAS = decimal.Parse( row["NUMERO_CUOTAS_VENCIDAS"].ToString());
						if (row["V_CUOTAS_VENCIDAS"] != System.DBNull.Value) V_CUOTAS_VENCIDAS = decimal.Parse( row["V_CUOTAS_VENCIDAS"].ToString());
						if (row["CUOTAS_CANCELADAS"] != System.DBNull.Value) CUOTAS_CANCELADAS = int.Parse( row["CUOTAS_CANCELADAS"].ToString());

						string sSQLAct = string.Format("Exec GrabaHistorialReporteComparaCarteraPichi '{0}', '{1}', '{2}', {3}, {4}, {5},{6},{7}, {8}, {9}, {10}, {11}, {12} ", 
							OPERACION	,
						IDENTIFICACION	,
							CLIENTE	,
							MONTO_FINANCIADO	,
							SALDO_CAPITAL_VENCIDO	,
							SALDO_CAPITAL_X_VENCER ,
							SALDO_CAPITAL_NODEV ,	
							INTERES_X_PAGAR	,
							MORA_X_PAGAR	,
							DIAS_MORA	,
							NUMERO_CUOTAS_VENCIDAS	,
							V_CUOTAS_VENCIDAS	,
							CUOTAS_CANCELADAS		);						
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
						iCont++;			
					}

					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;	
				return;										
			}
			#endregion Sube archivo
		}
	}
}
