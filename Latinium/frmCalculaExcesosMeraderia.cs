using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCalculaExcesosMeraderia.
	/// </summary>
	public class frmCalculaExcesosMeraderia : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridExcesoMeraderia;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTEM;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTVC;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTDM;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnCalcular;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTransferir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDeficitCategorias;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridExcesoCategorias;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnContinuar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.SaveFileDialog saveFileDialog2;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCalculaExcesosMeraderia()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		DateTime dtFecha = DateTime.Today;
		int iRangoDias = 0;
		int iDiasAbastecimiento = 0;
		bool bCalcular = false;
		int iDiasDeficit = 0;

		public frmCalculaExcesosMeraderia(DateTime DTFecha, int IRangoDias, int IDiasAbastecimiento, int IDiasDeficit, bool BCalcular)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			dtFecha = DTFecha;
			iRangoDias = IRangoDias;
			iDiasAbastecimiento = IDiasAbastecimiento;
			iDiasDeficit = IDiasDeficit;
			bCalcular = BCalcular;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadTransferido");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasUltTransf");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExistenciaAnterior");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaDiaria");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Requerido");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferido");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Requerimiento");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exceso");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exceso", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exceso", 12, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadTransferido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasUltTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ExistenciaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Requerido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Requerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Exceso");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasExcesosMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exceso");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exceso", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exceso", 5, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 6, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferenciasExcesosMercaderia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Exceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Distribucion");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deficit");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exceso");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EBI");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NEBI");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CNEBI");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deficit", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deficit", 2, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exceso", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exceso", 3, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "EBI", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "EBI", 4, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CNEBI", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CNEBI", 6, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deficit");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Exceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EBI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NEBI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CNEBI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Distribucion");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasDeficitMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deficit");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deficit", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deficit", 5, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 6, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferenciasDeficitMercaderia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deficit");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Distribucion");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasEDMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasExcesosMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasDeficitMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaExceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaDeficit");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 9, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferenciasEDMercaderia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferenciasExcesosMercaderia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransferenciasDeficitMercaderia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaExceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaDeficit");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasDeficitMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deficit");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deficit", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deficit", 5, true);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 6, true);
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransferenciasExcesosMercaderia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exceso");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exceso", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exceso", 5, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 6, true);
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCalculaExcesosMeraderia));
			this.uGridExcesoMeraderia = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTEM = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTVC = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTDM = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.uGridTransferir = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDeficitCategorias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridExcesoCategorias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnContinuar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
			this.btnExportar = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter2 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.uGridExcesoMeraderia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTEM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTVC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTDM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferir)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDeficitCategorias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridExcesoCategorias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridExcesoMeraderia
			// 
			this.uGridExcesoMeraderia.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridExcesoMeraderia.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridExcesoMeraderia.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridExcesoMeraderia.DisplayLayout.Appearance = appearance1;
			this.uGridExcesoMeraderia.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 62;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Articulo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 183;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "C Transferido";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.Caption = "D. U. T.";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 60;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.Caption = "E Anterior";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 65;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Header.Caption = "Venta Diaria";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Header.Caption = "Necesidad";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 72;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Header.Caption = "En Transito";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 60;
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
			ultraGridBand1.Header.Caption = "Calculo";
			ultraGridBand1.HeaderVisible = true;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.uGridExcesoMeraderia.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridExcesoMeraderia.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.uGridExcesoMeraderia.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridExcesoMeraderia.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridExcesoMeraderia.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.uGridExcesoMeraderia.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoMeraderia.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoMeraderia.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoMeraderia.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridExcesoMeraderia.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridExcesoMeraderia.Location = new System.Drawing.Point(0, 40);
			this.uGridExcesoMeraderia.Name = "uGridExcesoMeraderia";
			this.uGridExcesoMeraderia.Size = new System.Drawing.Size(986, 160);
			this.uGridExcesoMeraderia.TabIndex = 185;
			this.uGridExcesoMeraderia.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridExcesoMeraderia_BeforeSelectChange);
			this.uGridExcesoMeraderia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridExcesoMeraderia_InitializeLayout);
			this.uGridExcesoMeraderia.Click += new System.EventHandler(this.uGridExcesoMeraderia_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
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
			// uGridTEM
			// 
			this.uGridTEM.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTEM.DataSource = this.ultraDataSource2;
			appearance19.BackColor = System.Drawing.Color.White;
			this.uGridTEM.DisplayLayout.Appearance = appearance19;
			this.uGridTEM.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 129;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 131;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Articulo";
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 167;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 62;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Width = 42;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance20;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Width = 54;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance21;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Width = 57;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Width = 89;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand2.Header.Caption = "Exceso";
			ultraGridBand2.HeaderVisible = true;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance22;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance23;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridTEM.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTEM.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridTEM.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridTEM.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTEM.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.uGridTEM.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTEM.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTEM.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTEM.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridTEM.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTEM.Location = new System.Drawing.Point(8, 208);
			this.uGridTEM.Name = "uGridTEM";
			this.uGridTEM.Size = new System.Drawing.Size(447, 136);
			this.uGridTEM.TabIndex = 186;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// uGridTVC
			// 
			this.uGridTVC.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTVC.DataSource = this.ultraDataSource4;
			appearance30.BackColor = System.Drawing.Color.White;
			this.uGridTVC.DisplayLayout.Appearance = appearance30;
			this.uGridTVC.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Articulo";
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 174;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance31;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Width = 58;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance32;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Width = 59;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance33;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Width = 59;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance34;
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance35;
			ultraGridColumn28.Header.Caption = "C N BI";
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Width = 60;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Width = 86;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			ultraGridBand3.Header.Caption = "";
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance36;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance37;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance38;
			summarySettings6.DisplayFormat = "{0: #,##0}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance39;
			summarySettings7.DisplayFormat = "{0: #,##0}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridTVC.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTVC.DisplayLayout.Override.ActiveRowAppearance = appearance40;
			this.uGridTVC.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance41.BackColor = System.Drawing.Color.Transparent;
			this.uGridTVC.DisplayLayout.Override.CardAreaAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 8F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTVC.DisplayLayout.Override.HeaderAppearance = appearance42;
			this.uGridTVC.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance43.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance43.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTVC.DisplayLayout.Override.RowAlternateAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTVC.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTVC.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.uGridTVC.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTVC.Location = new System.Drawing.Point(8, 496);
			this.uGridTVC.Name = "uGridTVC";
			this.uGridTVC.Size = new System.Drawing.Size(517, 192);
			this.uGridTVC.TabIndex = 188;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(bool);
			ultraDataColumn27.DefaultValue = false;
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// uGridTDM
			// 
			this.uGridTDM.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTDM.DataSource = this.ultraDataSource3;
			appearance46.BackColor = System.Drawing.Color.White;
			this.uGridTDM.DisplayLayout.Appearance = appearance46;
			this.uGridTDM.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 123;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "Articulo";
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 168;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 62;
			ultraGridColumn34.Header.VisiblePosition = 5;
			ultraGridColumn34.Width = 43;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance47;
			ultraGridColumn35.Header.VisiblePosition = 6;
			ultraGridColumn35.Width = 54;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance48;
			ultraGridColumn36.Header.VisiblePosition = 7;
			ultraGridColumn36.Width = 57;
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 87;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			ultraGridBand4.Header.Caption = "Deficit";
			ultraGridBand4.HeaderVisible = true;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance49;
			summarySettings8.DisplayFormat = "{0: #,##0}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance50;
			summarySettings9.DisplayFormat = "{0: #,##0}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings8,
																																															summarySettings9});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridTDM.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance51.ForeColor = System.Drawing.Color.Black;
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTDM.DisplayLayout.Override.ActiveRowAppearance = appearance51;
			this.uGridTDM.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance52.BackColor = System.Drawing.Color.Transparent;
			this.uGridTDM.DisplayLayout.Override.CardAreaAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance53.FontData.BoldAsString = "True";
			appearance53.FontData.Name = "Arial";
			appearance53.FontData.SizeInPoints = 8F;
			appearance53.ForeColor = System.Drawing.Color.White;
			appearance53.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTDM.DisplayLayout.Override.HeaderAppearance = appearance53;
			this.uGridTDM.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance54.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance54.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTDM.DisplayLayout.Override.RowAlternateAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTDM.DisplayLayout.Override.RowSelectorAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTDM.DisplayLayout.Override.SelectedRowAppearance = appearance56;
			this.uGridTDM.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTDM.Location = new System.Drawing.Point(547, 208);
			this.uGridTDM.Name = "uGridTDM";
			this.uGridTDM.Size = new System.Drawing.Size(447, 136);
			this.uGridTDM.TabIndex = 187;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
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
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.ImageIndex = 62;
			this.btnCalcular.Location = new System.Drawing.Point(728, 8);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(128, 23);
			this.btnCalcular.TabIndex = 203;
			this.btnCalcular.Text = "Calcular Transferencia";
			this.btnCalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// uGridTransferir
			// 
			this.uGridTransferir.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTransferir.DataSource = this.ultraDataSource5;
			appearance57.BackColor = System.Drawing.Color.White;
			this.uGridTransferir.DisplayLayout.Appearance = appearance57;
			this.uGridTransferir.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 73;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 85;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 3;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 80;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn42.Header.VisiblePosition = 4;
			ultraGridColumn42.Width = 218;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn43.Header.VisiblePosition = 5;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 39;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 6;
			ultraGridColumn44.Width = 75;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 7;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 37;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.VisiblePosition = 8;
			ultraGridColumn46.Width = 75;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance58;
			ultraGridColumn47.Header.VisiblePosition = 9;
			ultraGridColumn47.Width = 75;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47});
			ultraGridBand5.Header.Caption = "";
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance59;
			summarySettings10.DisplayFormat = "{0: #,##0}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings10});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridTransferir.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance60.ForeColor = System.Drawing.Color.Black;
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransferir.DisplayLayout.Override.ActiveRowAppearance = appearance60;
			this.uGridTransferir.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance61.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransferir.DisplayLayout.Override.CardAreaAppearance = appearance61;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance62.FontData.BoldAsString = "True";
			appearance62.FontData.Name = "Arial";
			appearance62.FontData.SizeInPoints = 8F;
			appearance62.ForeColor = System.Drawing.Color.White;
			appearance62.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransferir.DisplayLayout.Override.HeaderAppearance = appearance62;
			this.uGridTransferir.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance63.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance63.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance63.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferir.DisplayLayout.Override.RowAlternateAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferir.DisplayLayout.Override.RowSelectorAppearance = appearance64;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferir.DisplayLayout.Override.SelectedRowAppearance = appearance65;
			this.uGridTransferir.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransferir.Location = new System.Drawing.Point(536, 496);
			this.uGridTransferir.Name = "uGridTransferir";
			this.uGridTransferir.Size = new System.Drawing.Size(464, 192);
			this.uGridTransferir.TabIndex = 204;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47});
			// 
			// uGridDeficitCategorias
			// 
			this.uGridDeficitCategorias.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDeficitCategorias.DataSource = this.ultraDataSource3;
			appearance66.BackColor = System.Drawing.Color.White;
			this.uGridDeficitCategorias.DisplayLayout.Appearance = appearance66;
			this.uGridDeficitCategorias.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 123;
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.Caption = "Articulo";
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Width = 169;
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 62;
			ultraGridColumn52.Header.VisiblePosition = 5;
			ultraGridColumn52.Width = 41;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance67;
			ultraGridColumn53.Header.VisiblePosition = 6;
			ultraGridColumn53.Width = 54;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance68;
			ultraGridColumn54.Header.VisiblePosition = 7;
			ultraGridColumn54.Width = 59;
			ultraGridColumn55.Header.VisiblePosition = 4;
			ultraGridColumn55.Width = 87;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			ultraGridBand6.Header.Caption = "Deficit";
			ultraGridBand6.HeaderVisible = true;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance69;
			summarySettings11.DisplayFormat = "{0: #,##0}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance70;
			summarySettings12.DisplayFormat = "{0: #,##0}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings11,
																																															summarySettings12});
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridDeficitCategorias.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance71.ForeColor = System.Drawing.Color.Black;
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDeficitCategorias.DisplayLayout.Override.ActiveRowAppearance = appearance71;
			this.uGridDeficitCategorias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance72.BackColor = System.Drawing.Color.Transparent;
			this.uGridDeficitCategorias.DisplayLayout.Override.CardAreaAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance73.FontData.BoldAsString = "True";
			appearance73.FontData.Name = "Arial";
			appearance73.FontData.SizeInPoints = 8F;
			appearance73.ForeColor = System.Drawing.Color.White;
			appearance73.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDeficitCategorias.DisplayLayout.Override.HeaderAppearance = appearance73;
			this.uGridDeficitCategorias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance74.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance74.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDeficitCategorias.DisplayLayout.Override.RowAlternateAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDeficitCategorias.DisplayLayout.Override.RowSelectorAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDeficitCategorias.DisplayLayout.Override.SelectedRowAppearance = appearance76;
			this.uGridDeficitCategorias.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDeficitCategorias.Location = new System.Drawing.Point(547, 352);
			this.uGridDeficitCategorias.Name = "uGridDeficitCategorias";
			this.uGridDeficitCategorias.Size = new System.Drawing.Size(448, 136);
			this.uGridDeficitCategorias.TabIndex = 206;
			// 
			// uGridExcesoCategorias
			// 
			this.uGridExcesoCategorias.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridExcesoCategorias.DataSource = this.ultraDataSource2;
			appearance77.BackColor = System.Drawing.Color.White;
			this.uGridExcesoCategorias.DisplayLayout.Appearance = appearance77;
			this.uGridExcesoCategorias.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 129;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 131;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn58.Header.Caption = "Articulo";
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Width = 167;
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 62;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 5;
			ultraGridColumn60.Width = 42;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance78;
			ultraGridColumn61.Header.VisiblePosition = 6;
			ultraGridColumn61.Width = 54;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance79;
			ultraGridColumn62.Header.VisiblePosition = 7;
			ultraGridColumn62.Width = 57;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn63.Header.VisiblePosition = 4;
			ultraGridColumn63.Width = 89;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63});
			ultraGridBand7.Header.Caption = "Exceso";
			ultraGridBand7.HeaderVisible = true;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance80;
			summarySettings13.DisplayFormat = "{0: #,##0}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance81;
			summarySettings14.DisplayFormat = "{0: #,##0}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand7.SummaryFooterCaption = "";
			this.uGridExcesoCategorias.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance82.ForeColor = System.Drawing.Color.Black;
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridExcesoCategorias.DisplayLayout.Override.ActiveRowAppearance = appearance82;
			this.uGridExcesoCategorias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance83.BackColor = System.Drawing.Color.Transparent;
			this.uGridExcesoCategorias.DisplayLayout.Override.CardAreaAppearance = appearance83;
			appearance84.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance84.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance84.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance84.FontData.BoldAsString = "True";
			appearance84.FontData.Name = "Arial";
			appearance84.FontData.SizeInPoints = 8F;
			appearance84.ForeColor = System.Drawing.Color.White;
			appearance84.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridExcesoCategorias.DisplayLayout.Override.HeaderAppearance = appearance84;
			this.uGridExcesoCategorias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance85.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance85.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoCategorias.DisplayLayout.Override.RowAlternateAppearance = appearance85;
			appearance86.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance86.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoCategorias.DisplayLayout.Override.RowSelectorAppearance = appearance86;
			appearance87.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance87.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance87.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExcesoCategorias.DisplayLayout.Override.SelectedRowAppearance = appearance87;
			this.uGridExcesoCategorias.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridExcesoCategorias.Location = new System.Drawing.Point(8, 352);
			this.uGridExcesoCategorias.Name = "uGridExcesoCategorias";
			this.uGridExcesoCategorias.Size = new System.Drawing.Size(447, 136);
			this.uGridExcesoCategorias.TabIndex = 205;
			// 
			// btnContinuar
			// 
			this.btnContinuar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContinuar.CausesValidation = false;
			this.btnContinuar.Enabled = false;
			this.btnContinuar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContinuar.ImageIndex = 62;
			this.btnContinuar.Location = new System.Drawing.Point(864, 8);
			this.btnContinuar.Name = "btnContinuar";
			this.btnContinuar.Size = new System.Drawing.Size(128, 23);
			this.btnContinuar.TabIndex = 207;
			this.btnContinuar.Text = "Continuar";
			this.btnContinuar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContinuar.Click += new System.EventHandler(this.btnContinuar_Click);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.ImageIndex = 62;
			this.btnVer.Location = new System.Drawing.Point(384, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(56, 23);
			this.btnVer.TabIndex = 215;
			this.btnVer.Text = "Ver";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 214;
			this.label1.Text = "Articulo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance88;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 8);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(304, 22);
			this.txtBuscar.TabIndex = 213;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(632, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 728;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmCalculaExcesosMeraderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1002, 696);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnContinuar);
			this.Controls.Add(this.uGridDeficitCategorias);
			this.Controls.Add(this.uGridExcesoCategorias);
			this.Controls.Add(this.uGridTransferir);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.uGridTVC);
			this.Controls.Add(this.uGridTDM);
			this.Controls.Add(this.uGridTEM);
			this.Controls.Add(this.uGridExcesoMeraderia);
			this.Controls.Add(this.btnExportar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCalculaExcesosMeraderia";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Calcula Excesos Meraderia";
			this.Load += new System.EventHandler(this.frmCalculaExcesosMeraderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridExcesoMeraderia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTEM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTVC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTDM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferir)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDeficitCategorias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridExcesoCategorias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCalculaExcesosMeraderia_Load(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			if (bCalcular) this.uGridExcesoMeraderia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec CalculaExistenciaEnExceso '{0}', {1}, {2}, {3}", dtFecha.ToString("yyyyMMdd"), iRangoDias, iDiasAbastecimiento, iDiasDeficit));
			if (!bCalcular) this.uGridExcesoMeraderia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 0, '', 0"));

			this.uGridTEM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 1, '', 0");
			this.uGridTDM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 2, '', 0");
			this.uGridExcesoCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 3, '', 0");
			this.uGridDeficitCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 4, '', 0");
			this.uGridTVC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 5, '', 0");
			this.uGridTransferir.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 6, '', 0");
			this.Cursor = Cursors.Default;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			if (this.txtBuscar.Text.ToString().Length > 0)
			{
				this.txtBuscar.Text = "";
				this.btnVer_Click(sender, e);
			}

			#region Distribucion Por Categorias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTEMC in this.uGridExcesoCategorias.Rows.All)
			{
				bool bExcesos = true;
				int iContArtDeficit = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From TransferenciasDeficitMercaderia Where idArticulo = {0} And Distribucion = 2 And Saldo > 0", (int)drTEMC.Cells["idArticulo"].Value));

				if (iContArtDeficit > 0 && (int)drTEMC.Cells["Saldo"].Value > 0)
				{
					int iPeso = Funcion.iEscalarSQL(cdsSeteoF, "Select Top 1 Peso From TipoDeAlmacen Order By Peso Desc");

					for (int i = iPeso; i >= 0; i--)
					{
						//int iBodegaDeficit = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @Bodega Int If Exists(Select Bodega From Bodega b Inner Join TipoDeAlmacen t On t.idTipoDeAlmacen = b.idTipoDeAlmacen Where t.Peso = {0}) Select @Bodega = Bodega From Bodega b Inner Join TipoDeAlmacen t On t.idTipoDeAlmacen = b.idTipoDeAlmacen Where t.Peso = {0} Order By Peso Desc Else Set @Bodega = 0 Select @Bodega", i));
            
						int iContBodegaPeso = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Bodega Where idTipoDeAlmacen In (Select idTipoDeAlmacen From TipoDeAlmacen Where Peso = {0})", i));

						for (int j = 1; j <= iContBodegaPeso; j++)
						{
							int iBodegaDeficit = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaBodegaTFMDC {0}, {1}", i, j));
							
							if (iBodegaDeficit > 0)
							{
								#region Recorre Tabla Excesos
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTDMC in this.uGridDeficitCategorias.Rows.All)
								{		
									if ((int)drTEMC.Cells["idArticulo"].Value == (int)drTDMC.Cells["idArticulo"].Value)
									{
										if (iBodegaDeficit == (int)drTDMC.Cells["NumBodega"].Value)
										{
											int iSaldoDeficit = (int)drTDMC.Cells["Saldo"].Value;
											int iSaldoExceso = (int)drTEMC.Cells["Saldo"].Value;
								
											int iTransfiere = 0;

											if (iSaldoDeficit > 0 && iSaldoExceso > 0)
											{
												#region Exceso Igual Que Deficit
												if (iSaldoExceso == iSaldoDeficit)
												{
													iTransfiere = iSaldoDeficit;
													iSaldoDeficit = 0;
													iSaldoExceso = 0;
													drTEMC.Cells["Saldo"].Value = iSaldoExceso;
													drTDMC.Cells["Saldo"].Value = iSaldoDeficit;
									
													string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 2",
														(int)drTDMC.Cells["idArticulo"].Value, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEMC.Cells["NumBodega"].Value, 
														(int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDMC.Cells["NumBodega"].Value, iTransfiere);
													Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
														iSaldoExceso, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
														iSaldoDeficit, (int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value));
													
													bExcesos = false;

													break;
												}
												#endregion Exceso Igual Que Deficit
										
												#region Exceso Mayor Que Deficit
												if (iSaldoExceso > iSaldoDeficit)
												{
													iTransfiere = iSaldoDeficit;
													iSaldoExceso = iSaldoExceso - iSaldoDeficit;
													iSaldoDeficit = 0;											
													drTEMC.Cells["Saldo"].Value = iSaldoExceso;
													drTDMC.Cells["Saldo"].Value = iSaldoDeficit;

													string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 2",
														(int)drTDMC.Cells["idArticulo"].Value, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEMC.Cells["NumBodega"].Value, 
														(int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDMC.Cells["NumBodega"].Value, iTransfiere);
													Funcion.EjecutaSQL(cdsSeteoF, sSQL);

													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
														iSaldoExceso, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
														iSaldoDeficit, (int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value));														
												}
												#endregion Exceso Mayor Que Deficit

												#region Exceso Menor Que Deficit
												if (iSaldoExceso < iSaldoDeficit)
												{
													iTransfiere = iSaldoExceso;
													iSaldoDeficit = iSaldoDeficit - iSaldoExceso;
													iSaldoExceso = 0;
													drTEMC.Cells["Saldo"].Value = iSaldoExceso;
													drTDMC.Cells["Saldo"].Value = iSaldoDeficit;

													string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 2",
														(int)drTDMC.Cells["idArticulo"].Value, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEMC.Cells["NumBodega"].Value, 
														(int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDMC.Cells["NumBodega"].Value, iTransfiere);
													Funcion.EjecutaSQL(cdsSeteoF, sSQL);

													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
														iSaldoExceso, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
														iSaldoDeficit, (int)drTDMC.Cells["idTransferenciasDeficitMercaderia"].Value));	
													
													break;

													bExcesos = false;													
												}
												#endregion Exceso Menor Que Deficit
											}
										}
									}
								}
								#endregion Recorre Tabla Excesos
							}
							if (!bExcesos) break;
						}
						if (!bExcesos) break;
					}
				}
			}
			#endregion Distribucion Por Categorias

			#region Distribucion Por Sectores
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTDM in this.uGridTDM.Rows.All)
			{					
				bool bDeficit = true;
				int iContArtExceso = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From TransferenciasExcesosMercaderia Where idArticulo = {0} And Distribucion = 1 And Saldo > 0", (int)drTDM.Cells["idArticulo"].Value));
				
				if (iContArtExceso > 0 && (int)drTDM.Cells["Saldo"].Value > 0)
				{
					int iBodegas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DistanciaEntreBodegas Where Bodega = {0}", (int)drTDM.Cells["NumBodega"].Value));
					
					for (int i = 1; i <= iBodegas; i++)
					{
						int iBodegaExceso = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Bodega1 From DistanciaEntreBodegas Where Distancia = {0} And Bodega = {1}", i, (int)drTDM.Cells["NumBodega"].Value));

						#region Recorre Tabla Excesos
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTEM in this.uGridTEM.Rows.All)
						{		
							if ((int)drTDM.Cells["idArticulo"].Value == (int)drTEM.Cells["idArticulo"].Value)
							{
								if (iBodegaExceso == (int)drTEM.Cells["NumBodega"].Value)
								{
									int iSaldoDeficit = (int)drTDM.Cells["Saldo"].Value;
									int iSaldoExceso = (int)drTEM.Cells["Saldo"].Value;
								
									int iTransfiere = 0;

									if (iSaldoDeficit > 0 && iSaldoExceso > 0)
									{
										#region Exceso Igual Que Deficit
										if (iSaldoExceso == iSaldoDeficit)
										{
											iTransfiere = iSaldoDeficit;
											iSaldoDeficit = 0;
											iSaldoExceso = 0;
											drTEM.Cells["Saldo"].Value = iSaldoExceso;
											drTDM.Cells["Saldo"].Value = iSaldoDeficit;
									
											string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 1",
												(int)drTDM.Cells["idArticulo"].Value, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEM.Cells["NumBodega"].Value, 
												(int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDM.Cells["NumBodega"].Value, iTransfiere);
											Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
												iSaldoExceso, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
												iSaldoDeficit, (int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value));	
											
											bDeficit = false;

											break;
										}
										#endregion Exceso Igual Que Deficit
										
										#region Exceso Mayor Que Deficit
										if (iSaldoExceso > iSaldoDeficit)
										{
											iTransfiere = iSaldoDeficit;
											iSaldoExceso = iSaldoExceso - iSaldoDeficit;
											iSaldoDeficit = 0;											
											drTEM.Cells["Saldo"].Value = iSaldoExceso;
											drTDM.Cells["Saldo"].Value = iSaldoDeficit;

											string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 1",
												(int)drTDM.Cells["idArticulo"].Value, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEM.Cells["NumBodega"].Value, 
												(int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDM.Cells["NumBodega"].Value, iTransfiere);
											Funcion.EjecutaSQL(cdsSeteoF, sSQL);

											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
												iSaldoExceso, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
												iSaldoDeficit, (int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value));	
											
											bDeficit = false;

											break;
										}
										#endregion Exceso Mayor Que Deficit

										#region Exceso Menor Que Deficit
										if (iSaldoExceso < iSaldoDeficit)
										{
											iTransfiere = iSaldoExceso;
											iSaldoDeficit = iSaldoDeficit - iSaldoExceso;
											iSaldoExceso = 0;
											drTEM.Cells["Saldo"].Value = iSaldoExceso;
											drTDM.Cells["Saldo"].Value = iSaldoDeficit;

											string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 1",
												(int)drTDM.Cells["idArticulo"].Value, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEM.Cells["NumBodega"].Value, 
												(int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value, (int)drTDM.Cells["NumBodega"].Value, iTransfiere);
											Funcion.EjecutaSQL(cdsSeteoF, sSQL);

											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = {0} Where idTransferenciasExcesosMercaderia = {1}",
												iSaldoExceso, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value));	
									
											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasDeficitMercaderia Set Saldo = {0} Where idTransferenciasDeficitMercaderia = {1}",
												iSaldoDeficit, (int)drTDM.Cells["idTransferenciasDeficitMercaderia"].Value));	
										}
										#endregion Exceso Menor Que Deficit
									}
								}
							}
						}
						#endregion Recorre Tabla Excesos
						if (!bDeficit)break;
					}					
				}
			}
			#endregion Distribucion Por Sectores			
			
			#region Distribucion a BI 
			#region Excesos Sectores
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTEM in this.uGridTEM.Rows.All)
			{	
				if ((int)drTEM.Cells["Saldo"].Value > 0)
				{
					string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 2",
						(int)drTEM.Cells["idArticulo"].Value, (int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEM.Cells["NumBodega"].Value, 
						0, 1, (int)drTEM.Cells["Saldo"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = 0 Where idTransferenciasExcesosMercaderia = {0}",
						(int)drTEM.Cells["idTransferenciasExcesosMercaderia"].Value));

					drTEM.Cells["Saldo"].Value = 0;
				}
			}
			#endregion Excesos Sectores

			#region Excesos Categorias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTEMC in this.uGridExcesoCategorias.Rows.All)
			{					
				if ((int)drTEMC.Cells["Saldo"].Value > 0)
				{
					string sSQL = string.Format("Exec GuardaTransferenciaEntreLocales {0}, {1}, {2}, {3}, {4}, {5}, 2",
						(int)drTEMC.Cells["idArticulo"].Value, (int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value, (int)drTEMC.Cells["NumBodega"].Value, 
						0, 1, (int)drTEMC.Cells["Saldo"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TransferenciasExcesosMercaderia Set Saldo = 0 Where idTransferenciasExcesosMercaderia = {0}",
						(int)drTEMC.Cells["idTransferenciasExcesosMercaderia"].Value));

					drTEMC.Cells["Saldo"].Value = 0;				
				}
			}
			#endregion Excesos Categorias			
			#endregion Distribucion a BI 

			this.Cursor = Cursors.Default;

			MessageBox.Show("Proceso Finalizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.uGridTransferir.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListasArticulosExcesosMercaderia 6, '', 0");

			this.btnContinuar.Enabled = true;
		}

		private void btnContinuar_Click(object sender, System.EventArgs e)
		{
			if (this.txtBuscar.Text.ToString().Length > 0)
			{
				this.txtBuscar.Text = "";
				this.btnVer_Click(sender, e);
			}

			this.DialogResult = DialogResult.OK;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridExcesoMeraderia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 0, '{0}', 0", this.txtBuscar.Text.ToString()));	
			this.uGridTEM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 1, '', 0"));
			this.uGridTDM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 2, '', 0"));
			this.uGridExcesoCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 3, '', 0"));
			this.uGridDeficitCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 4, '', 0"));
			this.uGridTVC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 5, '', 0"));
			this.uGridTransferir.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 6, '', 0"));
		}

		private void uGridExcesoMeraderia_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void uGridExcesoMeraderia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridExcesoMeraderia_Click(object sender, System.EventArgs e)
		{
			this.uGridTEM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 1, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
			this.uGridTDM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 2, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
			this.uGridExcesoCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 3, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
			this.uGridDeficitCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 4, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
			this.uGridTVC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 5, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
			this.uGridTransferir.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListasArticulosExcesosMercaderia 6, '', {0}", (int)this.uGridExcesoMeraderia.ActiveRow.Cells["idArticulo"].Value));
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridExcesoMeraderia.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{
					ultraGridExcelExporter1.Export(this.uGridExcesoMeraderia, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}		
	}
}

