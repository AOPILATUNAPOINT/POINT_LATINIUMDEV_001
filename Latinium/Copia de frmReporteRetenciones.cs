using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteRetenciones.
	/// </summary>
	public class frmReporteRetenciones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcRetenciones;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl upcEmitidas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl upcRecibidas;
		public DevExpress.XtraEditors.SimpleButton btnEmitidasExportar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEmitidasHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEmitidasDesde;
		private DevExpress.XtraEditors.SimpleButton btnEmitidasConsultar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtRecibidasDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtRecibidasHasta;
		private DevExpress.XtraEditors.SimpleButton btnRecibidasConsultar;
		private DevExpress.XtraEditors.SimpleButton btnExportarRecibidas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEmitidas;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRecibidas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		public DevExpress.XtraEditors.SimpleButton simpleButton1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHastaDE;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesdeDE;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lbl3;
		private System.Windows.Forms.Label lbl1;
		private System.Windows.Forms.Label lbl2;
		private DevExpress.XtraEditors.SimpleButton btnImprimir;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbIVA;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteRetenciones()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BI_IVA");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BI_Renta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 7, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Renta", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Renta", 10, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BI_IVA", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BI_IVA", 8, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BI_Renta", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BI_Renta", 9, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BI_IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BI_Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteRetenciones));
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BI_IVA");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BI_Renta");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 8, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Renta", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Renta", 10, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BI_IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BI_Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloSRI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoRenta");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetalleRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BIRenta");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoRenta");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoIVA");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetalleIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BI_IVA");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoIVA");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BIRenta", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BIRenta", 8, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoRenta", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoRenta", 9, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BI_IVA", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BI_IVA", 12, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoIVA", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoIVA", 13, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetalleRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BIRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetalleIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BI_IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoIVA");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.upcEmitidas = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lbl1 = new System.Windows.Forms.Label();
			this.uGridEmitidas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnEmitidasExportar = new DevExpress.XtraEditors.SimpleButton();
			this.dtEmitidasHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtEmitidasDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btnEmitidasConsultar = new DevExpress.XtraEditors.SimpleButton();
			this.upcRecibidas = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lbl2 = new System.Windows.Forms.Label();
			this.uGridRecibidas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtRecibidasDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtRecibidasHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnRecibidasConsultar = new DevExpress.XtraEditors.SimpleButton();
			this.label2 = new System.Windows.Forms.Label();
			this.btnExportarRecibidas = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbIVA = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.label46 = new System.Windows.Forms.Label();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.lbl3 = new System.Windows.Forms.Label();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.dtHastaDE = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesdeDE = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utcRetenciones = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.upcEmitidas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridEmitidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmitidasHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmitidasDesde)).BeginInit();
			this.upcRecibidas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecibidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibidasDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibidasHasta)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHastaDE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesdeDE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcRetenciones)).BeginInit();
			this.utcRetenciones.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// upcEmitidas
			// 
			this.upcEmitidas.Controls.Add(this.lbl1);
			this.upcEmitidas.Controls.Add(this.uGridEmitidas);
			this.upcEmitidas.Controls.Add(this.btnEmitidasExportar);
			this.upcEmitidas.Controls.Add(this.dtEmitidasHasta);
			this.upcEmitidas.Controls.Add(this.dtEmitidasDesde);
			this.upcEmitidas.Controls.Add(this.label3);
			this.upcEmitidas.Controls.Add(this.label4);
			this.upcEmitidas.Controls.Add(this.btnEmitidasConsultar);
			this.upcEmitidas.Location = new System.Drawing.Point(2, 21);
			this.upcEmitidas.Name = "upcEmitidas";
			this.upcEmitidas.Size = new System.Drawing.Size(1048, 529);
			// 
			// lbl1
			// 
			this.lbl1.AutoSize = true;
			this.lbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl1.ForeColor = System.Drawing.Color.Firebrick;
			this.lbl1.Location = new System.Drawing.Point(680, 5);
			this.lbl1.Name = "lbl1";
			this.lbl1.Size = new System.Drawing.Size(0, 26);
			this.lbl1.TabIndex = 643;
			// 
			// uGridEmitidas
			// 
			this.uGridEmitidas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEmitidas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEmitidas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEmitidas.DisplayLayout.Appearance = appearance1;
			this.uGridEmitidas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 65;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Proveedor";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 250;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 65;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance6;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 75;
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
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnHeaderAppearance = appearance7;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance10;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance11;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance12;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridEmitidas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridEmitidas.DisplayLayout.GroupByBox.Hidden = true;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEmitidas.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridEmitidas.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEmitidas.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEmitidas.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEmitidas.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEmitidas.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridEmitidas.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridEmitidas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEmitidas.Location = new System.Drawing.Point(6, 40);
			this.uGridEmitidas.Name = "uGridEmitidas";
			this.uGridEmitidas.Size = new System.Drawing.Size(1036, 488);
			this.uGridEmitidas.TabIndex = 630;
			this.uGridEmitidas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridEmitidas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
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
			// btnEmitidasExportar
			// 
			this.btnEmitidasExportar.Enabled = false;
			this.btnEmitidasExportar.Location = new System.Drawing.Point(520, 8);
			this.btnEmitidasExportar.Name = "btnEmitidasExportar";
			this.btnEmitidasExportar.Size = new System.Drawing.Size(80, 21);
			this.btnEmitidasExportar.TabIndex = 629;
			this.btnEmitidasExportar.Text = "&Exportar";
			this.btnEmitidasExportar.Click += new System.EventHandler(this.btnEmitidasExportar_Click);
			// 
			// dtEmitidasHasta
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtEmitidasHasta.Appearance = appearance19;
			this.dtEmitidasHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtEmitidasHasta.DateButtons.Add(dateButton1);
			this.dtEmitidasHasta.Format = "dd/MM/yyyy";
			this.dtEmitidasHasta.Location = new System.Drawing.Point(256, 8);
			this.dtEmitidasHasta.Name = "dtEmitidasHasta";
			this.dtEmitidasHasta.NonAutoSizeHeight = 23;
			this.dtEmitidasHasta.Size = new System.Drawing.Size(112, 21);
			this.dtEmitidasHasta.SpinButtonsVisible = true;
			this.dtEmitidasHasta.TabIndex = 327;
			this.dtEmitidasHasta.Value = ((object)(resources.GetObject("dtEmitidasHasta.Value")));
			// 
			// dtEmitidasDesde
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtEmitidasDesde.Appearance = appearance20;
			this.dtEmitidasDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtEmitidasDesde.DateButtons.Add(dateButton2);
			this.dtEmitidasDesde.Format = "dd/MM/yyyy";
			this.dtEmitidasDesde.Location = new System.Drawing.Point(64, 8);
			this.dtEmitidasDesde.Name = "dtEmitidasDesde";
			this.dtEmitidasDesde.NonAutoSizeHeight = 23;
			this.dtEmitidasDesde.Size = new System.Drawing.Size(112, 21);
			this.dtEmitidasDesde.SpinButtonsVisible = true;
			this.dtEmitidasDesde.TabIndex = 326;
			this.dtEmitidasDesde.Value = ((object)(resources.GetObject("dtEmitidasDesde.Value")));
			this.dtEmitidasDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtMayorDesde_BeforeDropDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(200, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 333;
			this.label3.Text = "Hasta:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 332;
			this.label4.Text = "Desde:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnEmitidasConsultar
			// 
			this.btnEmitidasConsultar.Location = new System.Drawing.Point(400, 8);
			this.btnEmitidasConsultar.Name = "btnEmitidasConsultar";
			this.btnEmitidasConsultar.Size = new System.Drawing.Size(80, 21);
			this.btnEmitidasConsultar.TabIndex = 331;
			this.btnEmitidasConsultar.Text = "Generar";
			this.btnEmitidasConsultar.Click += new System.EventHandler(this.btnEmitidasConsultar_Click);
			// 
			// upcRecibidas
			// 
			this.upcRecibidas.Controls.Add(this.lbl2);
			this.upcRecibidas.Controls.Add(this.uGridRecibidas);
			this.upcRecibidas.Controls.Add(this.label1);
			this.upcRecibidas.Controls.Add(this.dtRecibidasDesde);
			this.upcRecibidas.Controls.Add(this.dtRecibidasHasta);
			this.upcRecibidas.Controls.Add(this.btnRecibidasConsultar);
			this.upcRecibidas.Controls.Add(this.label2);
			this.upcRecibidas.Controls.Add(this.btnExportarRecibidas);
			this.upcRecibidas.Location = new System.Drawing.Point(-10000, -10000);
			this.upcRecibidas.Name = "upcRecibidas";
			this.upcRecibidas.Size = new System.Drawing.Size(1048, 529);
			this.upcRecibidas.Paint += new System.Windows.Forms.PaintEventHandler(this.upcRecibidas_Paint);
			// 
			// lbl2
			// 
			this.lbl2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lbl2.AutoSize = true;
			this.lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl2.ForeColor = System.Drawing.Color.Firebrick;
			this.lbl2.Location = new System.Drawing.Point(476, 8);
			this.lbl2.Name = "lbl2";
			this.lbl2.Size = new System.Drawing.Size(0, 26);
			this.lbl2.TabIndex = 643;
			// 
			// uGridRecibidas
			// 
			this.uGridRecibidas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRecibidas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRecibidas.DataSource = this.ultraDataSource2;
			appearance21.BackColor = System.Drawing.Color.White;
			this.uGridRecibidas.DisplayLayout.Appearance = appearance21;
			this.uGridRecibidas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 103;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 65;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Width = 250;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Width = 65;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance22;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Width = 75;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance23;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 7;
			ultraGridColumn21.Width = 75;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance24;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 11;
			ultraGridColumn22.Width = 75;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance25;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance26;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Width = 75;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24});
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance27;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance28;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance29;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8});
			this.uGridRecibidas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridRecibidas.DisplayLayout.GroupByBox.Hidden = true;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRecibidas.DisplayLayout.Override.ActiveRowAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.uGridRecibidas.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 8F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRecibidas.DisplayLayout.Override.HeaderAppearance = appearance32;
			appearance33.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance33.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecibidas.DisplayLayout.Override.RowAlternateAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecibidas.DisplayLayout.Override.RowSelectorAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecibidas.DisplayLayout.Override.SelectedRowAppearance = appearance35;
			this.uGridRecibidas.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridRecibidas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRecibidas.Location = new System.Drawing.Point(6, 40);
			this.uGridRecibidas.Name = "uGridRecibidas";
			this.uGridRecibidas.Size = new System.Drawing.Size(1036, 488);
			this.uGridRecibidas.TabIndex = 631;
			this.uGridRecibidas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridRecibidas_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn15.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn21.DataType = typeof(System.Double);
			ultraDataColumn21.DefaultValue = 0;
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn23.DataType = typeof(System.Double);
			ultraDataColumn23.DefaultValue = 0;
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn24.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 13);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 223;
			this.label1.Text = "Desde:";
			// 
			// dtRecibidasDesde
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtRecibidasDesde.Appearance = appearance36;
			this.dtRecibidasDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtRecibidasDesde.DateButtons.Add(dateButton3);
			this.dtRecibidasDesde.Format = "dd/MM/yyyy";
			this.dtRecibidasDesde.Location = new System.Drawing.Point(72, 11);
			this.dtRecibidasDesde.Name = "dtRecibidasDesde";
			this.dtRecibidasDesde.NonAutoSizeHeight = 23;
			this.dtRecibidasDesde.Size = new System.Drawing.Size(112, 21);
			this.dtRecibidasDesde.SpinButtonsVisible = true;
			this.dtRecibidasDesde.TabIndex = 225;
			this.dtRecibidasDesde.Value = ((object)(resources.GetObject("dtRecibidasDesde.Value")));
			// 
			// dtRecibidasHasta
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtRecibidasHasta.Appearance = appearance37;
			this.dtRecibidasHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtRecibidasHasta.DateButtons.Add(dateButton4);
			this.dtRecibidasHasta.Format = "dd/MM/yyyy";
			this.dtRecibidasHasta.Location = new System.Drawing.Point(288, 11);
			this.dtRecibidasHasta.Name = "dtRecibidasHasta";
			this.dtRecibidasHasta.NonAutoSizeHeight = 23;
			this.dtRecibidasHasta.Size = new System.Drawing.Size(112, 21);
			this.dtRecibidasHasta.SpinButtonsVisible = true;
			this.dtRecibidasHasta.TabIndex = 226;
			this.dtRecibidasHasta.Value = ((object)(resources.GetObject("dtRecibidasHasta.Value")));
			// 
			// btnRecibidasConsultar
			// 
			this.btnRecibidasConsultar.Location = new System.Drawing.Point(432, 11);
			this.btnRecibidasConsultar.Name = "btnRecibidasConsultar";
			this.btnRecibidasConsultar.Size = new System.Drawing.Size(80, 21);
			this.btnRecibidasConsultar.TabIndex = 218;
			this.btnRecibidasConsultar.Text = "Generar";
			this.btnRecibidasConsultar.Click += new System.EventHandler(this.btnRecibidasConsultar_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(224, 13);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 224;
			this.label2.Text = "Hasta:";
			// 
			// btnExportarRecibidas
			// 
			this.btnExportarRecibidas.Enabled = false;
			this.btnExportarRecibidas.Location = new System.Drawing.Point(552, 11);
			this.btnExportarRecibidas.Name = "btnExportarRecibidas";
			this.btnExportarRecibidas.Size = new System.Drawing.Size(80, 21);
			this.btnExportarRecibidas.TabIndex = 220;
			this.btnExportarRecibidas.Text = "&Exportar";
			this.btnExportarRecibidas.Click += new System.EventHandler(this.btnExportarRecibidas_Click);
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.cmbIVA);
			this.ultraTabPageControl1.Controls.Add(this.label7);
			this.ultraTabPageControl1.Controls.Add(this.label46);
			this.ultraTabPageControl1.Controls.Add(this.cmbArticuloSRI);
			this.ultraTabPageControl1.Controls.Add(this.btnImprimir);
			this.ultraTabPageControl1.Controls.Add(this.lbl3);
			this.ultraTabPageControl1.Controls.Add(this.simpleButton1);
			this.ultraTabPageControl1.Controls.Add(this.dtHastaDE);
			this.ultraTabPageControl1.Controls.Add(this.dtDesdeDE);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.label6);
			this.ultraTabPageControl1.Controls.Add(this.simpleButton2);
			this.ultraTabPageControl1.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1048, 529);
			// 
			// cmbIVA
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbIVA.Appearance = appearance38;
			this.cmbIVA.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbIVA.Cursor = System.Windows.Forms.Cursors.IBeam;
			this.cmbIVA.DataSource = this.ultraDataSource5;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 73;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 248;
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Width = 57;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbIVA.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbIVA.DisplayMember = "Detalle";
			this.cmbIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbIVA.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbIVA.Location = new System.Drawing.Point(1064, 8);
			this.cmbIVA.Name = "cmbIVA";
			this.cmbIVA.Size = new System.Drawing.Size(256, 21);
			this.cmbIVA.TabIndex = 654;
			this.cmbIVA.ValueMember = "idRetIva";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.Double);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			this.ultraDataSource5.Band.Key = "ArticuloSRI";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(1024, 11);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(32, 16);
			this.label7.TabIndex = 652;
			this.label7.Text = "I.V.A.";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(712, 11);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(34, 16);
			this.label46.TabIndex = 651;
			this.label46.Text = "Renta";
			this.label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArticuloSRI
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance39;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticuloSRI.DataSource = this.ultraDataSource4;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 113;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 75;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 250;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance40;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "%";
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Width = 50;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(760, 8);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(256, 21);
			this.cmbArticuloSRI.TabIndex = 650;
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(System.Double);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			this.ultraDataSource4.Band.Key = "ArticuloSRI";
			// 
			// btnImprimir
			// 
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Location = new System.Drawing.Point(488, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 21);
			this.btnImprimir.TabIndex = 649;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// lbl3
			// 
			this.lbl3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lbl3.AutoSize = true;
			this.lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl3.ForeColor = System.Drawing.Color.Firebrick;
			this.lbl3.Location = new System.Drawing.Point(540, 5);
			this.lbl3.Name = "lbl3";
			this.lbl3.Size = new System.Drawing.Size(0, 26);
			this.lbl3.TabIndex = 642;
			// 
			// simpleButton1
			// 
			this.simpleButton1.Enabled = false;
			this.simpleButton1.Location = new System.Drawing.Point(576, 8);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(80, 21);
			this.simpleButton1.TabIndex = 637;
			this.simpleButton1.Text = "&Exportar";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// dtHastaDE
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHastaDE.Appearance = appearance41;
			this.dtHastaDE.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtHastaDE.DateButtons.Add(dateButton5);
			this.dtHastaDE.Format = "dd/MM/yyyy";
			this.dtHastaDE.Location = new System.Drawing.Point(256, 8);
			this.dtHastaDE.Name = "dtHastaDE";
			this.dtHastaDE.NonAutoSizeHeight = 23;
			this.dtHastaDE.Size = new System.Drawing.Size(112, 21);
			this.dtHastaDE.SpinButtonsVisible = true;
			this.dtHastaDE.TabIndex = 633;
			this.dtHastaDE.Value = ((object)(resources.GetObject("dtHastaDE.Value")));
			// 
			// dtDesdeDE
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesdeDE.Appearance = appearance42;
			this.dtDesdeDE.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtDesdeDE.DateButtons.Add(dateButton6);
			this.dtDesdeDE.Format = "dd/MM/yyyy";
			this.dtDesdeDE.Location = new System.Drawing.Point(64, 8);
			this.dtDesdeDE.Name = "dtDesdeDE";
			this.dtDesdeDE.NonAutoSizeHeight = 23;
			this.dtDesdeDE.Size = new System.Drawing.Size(112, 21);
			this.dtDesdeDE.SpinButtonsVisible = true;
			this.dtDesdeDE.TabIndex = 632;
			this.dtDesdeDE.Value = ((object)(resources.GetObject("dtDesdeDE.Value")));
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(200, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 636;
			this.label5.Text = "Hasta:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 16);
			this.label6.TabIndex = 635;
			this.label6.Text = "Desde:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// simpleButton2
			// 
			this.simpleButton2.Location = new System.Drawing.Point(400, 8);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(80, 21);
			this.simpleButton2.TabIndex = 634;
			this.simpleButton2.Text = "Generar";
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance43.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance43;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Width = 72;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 51;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Width = 78;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 157;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 70;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.Caption = "Fecha Ret";
			ultraGridColumn38.Header.VisiblePosition = 5;
			ultraGridColumn38.Width = 51;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn39.CellAppearance = appearance44;
			ultraGridColumn39.Header.Caption = "Cod_Renta";
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridColumn39.Width = 54;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 7;
			ultraGridColumn40.Width = 102;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance45;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "BI Renta";
			ultraGridColumn41.Header.VisiblePosition = 8;
			ultraGridColumn41.Width = 54;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance46;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Ret_Renta";
			ultraGridColumn42.Header.VisiblePosition = 9;
			ultraGridColumn42.Width = 54;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn43.CellAppearance = appearance47;
			ultraGridColumn43.Format = "";
			ultraGridColumn43.Header.Caption = "Cod_IVA";
			ultraGridColumn43.Header.VisiblePosition = 10;
			ultraGridColumn43.Width = 54;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 11;
			ultraGridColumn44.Width = 102;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance48;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "BI IVA";
			ultraGridColumn45.Header.VisiblePosition = 12;
			ultraGridColumn45.Width = 54;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance49;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.Caption = "Ret_IVA";
			ultraGridColumn46.Header.VisiblePosition = 13;
			ultraGridColumn46.Width = 63;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance50;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance51;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance52;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance53;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance56.FontData.BoldAsString = "True";
			appearance56.FontData.Name = "Arial";
			appearance56.FontData.SizeInPoints = 8F;
			appearance56.ForeColor = System.Drawing.Color.White;
			appearance56.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance56;
			appearance57.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance57.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance59;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(6, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1036, 488);
			this.ultraGrid1.TabIndex = 631;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn34.DataType = typeof(System.DateTime);
			ultraDataColumn34.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn38.DataType = typeof(System.DateTime);
			ultraDataColumn38.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn41.DataType = typeof(System.Double);
			ultraDataColumn41.DefaultValue = 0;
			ultraDataColumn42.DataType = typeof(System.Double);
			ultraDataColumn42.DefaultValue = 0;
			ultraDataColumn45.DataType = typeof(System.Double);
			ultraDataColumn45.DefaultValue = 0;
			ultraDataColumn46.DataType = typeof(System.Double);
			ultraDataColumn46.DefaultValue = 0;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46});
			// 
			// utcRetenciones
			// 
			this.utcRetenciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcRetenciones.Appearance = appearance60;
			this.utcRetenciones.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcRetenciones.Controls.Add(this.upcEmitidas);
			this.utcRetenciones.Controls.Add(this.upcRecibidas);
			this.utcRetenciones.Controls.Add(this.ultraTabPageControl1);
			this.utcRetenciones.Location = new System.Drawing.Point(8, 8);
			this.utcRetenciones.Name = "utcRetenciones";
			this.utcRetenciones.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcRetenciones.Size = new System.Drawing.Size(1052, 552);
			this.utcRetenciones.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.PropertyPage2003;
			this.utcRetenciones.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.OfficeXPToolbarButton;
			this.utcRetenciones.TabIndex = 229;
			ultraTab1.TabPage = this.upcEmitidas;
			ultraTab1.Text = "Emitidas";
			ultraTab2.TabPage = this.upcRecibidas;
			ultraTab2.Text = "Recibidas";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Detalle Emitidas";
			ultraTab3.Visible = false;
			this.utcRetenciones.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3});
			this.utcRetenciones.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1048, 529);
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
			// frmReporteRetenciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1068, 566);
			this.Controls.Add(this.utcRetenciones);
			this.Name = "frmReporteRetenciones";
			this.Text = "Reporte de Retenciones";
			this.Load += new System.EventHandler(this.frmReporteRetenciones_Load);
			this.upcEmitidas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridEmitidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmitidasHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmitidasDesde)).EndInit();
			this.upcRecibidas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridRecibidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibidasDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibidasHasta)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHastaDE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesdeDE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcRetenciones)).EndInit();
			this.utcRetenciones.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void dtMayorDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnEmitidasConsultar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtEmitidasDesde.Value;
			DateTime dtHasta = (DateTime)this.dtEmitidasHasta.Value;

			string sSQL = string.Format("Exec ReporteRetenciones 1, '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.uGridEmitidas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lbl1.Text = this.uGridEmitidas.Rows.Count + " REGISTROS";
		}

		private void btnRecibidasConsultar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtRecibidasDesde.Value;
			DateTime dtHasta = (DateTime)this.dtRecibidasHasta.Value;

			string sSQL = string.Format("Exec ReporteRetenciones 0, '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.uGridRecibidas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lbl2.Text = this.uGridRecibidas.Rows.Count + " REGISTROS";
		}

		private Acceso miAcceso;

		private void frmReporteRetenciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso (cdsSeteoF,"200202");

			if (miAcceso.BExportar) this.btnEmitidasExportar.Enabled = true;
			if (miAcceso.Nuevo) this.btnExportarRecibidas.Enabled = true;
			if (miAcceso.Buscar) this.btnImprimir.Enabled = true;
			if (miAcceso.Editar) this.simpleButton1.Enabled = true;

			


			this.dtEmitidasDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtEmitidasHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtRecibidasDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtRecibidasHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtDesdeDE.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHastaDE.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.cmbArticuloSRI.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Detalle, Porcentaje From ArticuloSRI Order By Codigo");
			this.cmbIVA.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Codigo, Detalle, Porcentaje From RetencionIva Order By Codigo");
		}

		private void uGridEmitidas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void uGridRecibidas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void btnEmitidasExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridEmitidas.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.uGridEmitidas, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnExportarRecibidas_Click(object sender, System.EventArgs e)
		{
			if (this.uGridRecibidas.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.uGridRecibidas, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesdeDE.Value;
			DateTime dtHasta = (DateTime)this.dtHastaDE.Value;

			int idRenta = 0;
			int idIVA = 0;

			if (this.cmbArticuloSRI.ActiveRow != null) idRenta = (int)this.cmbArticuloSRI.Value;
			if (this.cmbIVA.ActiveRow != null) idIVA = (int)this.cmbIVA.Value;

			string sSQL = string.Format("Exec ReporteDetalleRetenciones '{0}', '{1}', {2}, {3}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), idRenta, idIVA);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);		
	
			this.lbl3.Text = this.ultraGrid1.Rows.Count + " REGISTROS";
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.ultraGrid1, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesdeDE.Value;
			DateTime dtHasta = (DateTime)this.dtHastaDE.Value;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = dtDesde;
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = dtHasta;
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);

			Reporte Reporte = new Reporte("RetencionesEmitidas", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void upcRecibidas_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}
	}
}
