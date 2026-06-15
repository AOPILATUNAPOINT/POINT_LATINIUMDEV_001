using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmComparaEstadosDeCuenta.
	/// </summary>
	public class frmComparaEstadosDeCuenta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridEntradaNoConciliados;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSalidasNoConciliados;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSalidaConciliados;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridEntradaConciliados;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridEntradaNoRegistrados;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSalidaNoRegistrados;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExportarEntradaSS;
		private System.Windows.Forms.Button btnExportarEntradaEC;
		private System.Windows.Forms.Button btnExportarSalidaEC;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnExportarSalidaSS;
		private System.Windows.Forms.Button btnConciliar;
		private System.Windows.Forms.Button btnProcesarModificacion;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmComparaEstadosDeCuenta()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		Infragistics.Win.UltraWinGrid.UltraGrid uGridEntradas;
		Infragistics.Win.UltraWinGrid.UltraGrid uGridSalidas;
		Infragistics.Win.UltraWinGrid.UltraGrid uGridECEntradas;
		Infragistics.Win.UltraWinGrid.UltraGrid uGridECSalidas;

		public frmComparaEstadosDeCuenta(Infragistics.Win.UltraWinGrid.UltraGrid UGridEntradas, Infragistics.Win.UltraWinGrid.UltraGrid UGridSalidas, Infragistics.Win.UltraWinGrid.UltraGrid UGridECEntradas, Infragistics.Win.UltraWinGrid.UltraGrid UGridECSalidas)
		{
			InitializeComponent();

			uGridEntradas	= UGridEntradas;
			uGridSalidas = UGridSalidas;
			uGridECEntradas	= UGridECEntradas;
			uGridECSalidas = UGridECSalidas;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 4, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 4, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			this.uGridEntradaNoConciliados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSalidasNoConciliados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSalidaConciliados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridEntradaConciliados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridEntradaNoRegistrados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSalidaNoRegistrados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid4 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnConciliar = new System.Windows.Forms.Button();
			this.btnExportarEntradaSS = new System.Windows.Forms.Button();
			this.btnExportarEntradaEC = new System.Windows.Forms.Button();
			this.btnExportarSalidaSS = new System.Windows.Forms.Button();
			this.btnExportarSalidaEC = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnProcesarModificacion = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaNoConciliados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidasNoConciliados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidaConciliados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaConciliados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaNoRegistrados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidaNoRegistrados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridEntradaNoConciliados
			// 
			this.uGridEntradaNoConciliados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntradaNoConciliados.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEntradaNoConciliados.DisplayLayout.Appearance = appearance1;
			this.uGridEntradaNoConciliados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 39;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 50;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 76;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Tipo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 77;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Numero";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 77;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 66;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 29;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "Movimientos De Entrada Solo En Sistema";
			ultraGridBand1.HeaderVisible = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridEntradaNoConciliados.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridEntradaNoConciliados.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridEntradaNoConciliados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntradaNoConciliados.Location = new System.Drawing.Point(360, 8);
			this.uGridEntradaNoConciliados.Name = "uGridEntradaNoConciliados";
			this.uGridEntradaNoConciliados.Size = new System.Drawing.Size(346, 298);
			this.uGridEntradaNoConciliados.TabIndex = 658;
			this.uGridEntradaNoConciliados.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntradaNoConciliados_AfterCellUpdate);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// uGridSalidasNoConciliados
			// 
			this.uGridSalidasNoConciliados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSalidasNoConciliados.DataSource = this.ultraDataSource4;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridSalidasNoConciliados.DisplayLayout.Appearance = appearance9;
			this.uGridSalidasNoConciliados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 39;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 50;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 76;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Tipo";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 77;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Numero";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 77;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance10;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Width = 66;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 29;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			ultraGridBand2.Header.Caption = "Movimientos de Salida Solo En Sistema";
			ultraGridBand2.HeaderVisible = true;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance11;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridSalidasNoConciliados.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8.5F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridSalidasNoConciliados.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridSalidasNoConciliados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalidasNoConciliados.Location = new System.Drawing.Point(360, 344);
			this.uGridSalidasNoConciliados.Name = "uGridSalidasNoConciliados";
			this.uGridSalidasNoConciliados.Size = new System.Drawing.Size(346, 298);
			this.uGridSalidasNoConciliados.TabIndex = 657;
			this.uGridSalidasNoConciliados.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalidasNoConciliados_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// uGridSalidaConciliados
			// 
			this.uGridSalidaConciliados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSalidaConciliados.DataSource = this.ultraDataSource3;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridSalidaConciliados.DisplayLayout.Appearance = appearance17;
			this.uGridSalidaConciliados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 26;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 33;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 52;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.Caption = "Tipo";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 98;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Numero";
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 98;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance18;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 77;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			ultraGridBand3.Header.Caption = "Movimientos de Salida Conciliados";
			ultraGridBand3.HeaderVisible = true;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance19;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridSalidaConciliados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalidaConciliados.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.uGridSalidaConciliados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSalidaConciliados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalidaConciliados.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8.25F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalidaConciliados.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.uGridSalidaConciliados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance23.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance23.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidaConciliados.DisplayLayout.Override.RowAlternateAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidaConciliados.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidaConciliados.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.uGridSalidaConciliados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalidaConciliados.Location = new System.Drawing.Point(8, 344);
			this.uGridSalidaConciliados.Name = "uGridSalidaConciliados";
			this.uGridSalidaConciliados.Size = new System.Drawing.Size(346, 296);
			this.uGridSalidaConciliados.TabIndex = 660;
			this.uGridSalidaConciliados.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridSalidaConciliados.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalidaConciliados_AfterCellUpdate);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// uGridEntradaConciliados
			// 
			this.uGridEntradaConciliados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntradaConciliados.DataSource = this.ultraDataSource1;
			appearance26.BackColor = System.Drawing.Color.White;
			this.uGridEntradaConciliados.DisplayLayout.Appearance = appearance26;
			this.uGridEntradaConciliados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 95;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 115;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 77;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.Caption = "Tipo";
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Width = 54;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Número";
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Width = 97;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance27;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Width = 97;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			ultraGridBand4.Header.Caption = "Movimientos de Entrada Conciliados";
			ultraGridBand4.HeaderVisible = true;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance28;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridEntradaConciliados.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntradaConciliados.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.uGridEntradaConciliados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntradaConciliados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntradaConciliados.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8.25F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntradaConciliados.DisplayLayout.Override.HeaderAppearance = appearance31;
			this.uGridEntradaConciliados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance32.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance32.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaConciliados.DisplayLayout.Override.RowAlternateAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaConciliados.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaConciliados.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridEntradaConciliados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntradaConciliados.Location = new System.Drawing.Point(7, 8);
			this.uGridEntradaConciliados.Name = "uGridEntradaConciliados";
			this.uGridEntradaConciliados.Size = new System.Drawing.Size(346, 296);
			this.uGridEntradaConciliados.TabIndex = 659;
			this.uGridEntradaConciliados.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridEntradaConciliados.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntradaConciliados_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// uGridEntradaNoRegistrados
			// 
			this.uGridEntradaNoRegistrados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntradaNoRegistrados.DataSource = this.ultraDataSource5;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridEntradaNoRegistrados.DisplayLayout.Appearance = appearance35;
			this.uGridEntradaNoRegistrados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 39;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Format = "dd/MM/yyyy";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 73;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Width = 77;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 77;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance36;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 67;
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			ultraGridBand5.Header.Caption = "Movimientos de Entrada Solo En Estado de Cuenta";
			ultraGridBand5.HeaderVisible = true;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance37;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings5});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridEntradaNoRegistrados.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.ActiveRowAppearance = appearance38;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance39.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.CardAreaAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance40.FontData.BoldAsString = "True";
			appearance40.FontData.Name = "Arial";
			appearance40.FontData.SizeInPoints = 8.5F;
			appearance40.ForeColor = System.Drawing.Color.White;
			appearance40.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.HeaderAppearance = appearance40;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.RowSelectorAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.SelectedRowAppearance = appearance42;
			this.uGridEntradaNoRegistrados.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridEntradaNoRegistrados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntradaNoRegistrados.Location = new System.Drawing.Point(712, 8);
			this.uGridEntradaNoRegistrados.Name = "uGridEntradaNoRegistrados";
			this.uGridEntradaNoRegistrados.Size = new System.Drawing.Size(346, 298);
			this.uGridEntradaNoRegistrados.TabIndex = 662;
			this.uGridEntradaNoRegistrados.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntradaNoRegistrados_CellChange);
			this.uGridEntradaNoRegistrados.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn32.DataType = typeof(bool);
			ultraDataColumn32.DefaultValue = false;
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// uGridSalidaNoRegistrados
			// 
			this.uGridSalidaNoRegistrados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSalidaNoRegistrados.DataSource = this.ultraDataSource6;
			appearance43.BackColor = System.Drawing.Color.White;
			this.uGridSalidaNoRegistrados.DisplayLayout.Appearance = appearance43;
			this.uGridSalidaNoRegistrados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 39;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 76;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Width = 81;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 77;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance44;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 66;
			ultraGridColumn38.Header.VisiblePosition = 5;
			ultraGridColumn38.Width = 25;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			ultraGridBand6.Header.Caption = "Movimientos de Salida Solo En Estado de Cuenta";
			ultraGridBand6.HeaderVisible = true;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance45;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6});
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridSalidaNoRegistrados.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance48.FontData.BoldAsString = "True";
			appearance48.FontData.Name = "Arial";
			appearance48.FontData.SizeInPoints = 8.5F;
			appearance48.ForeColor = System.Drawing.Color.White;
			appearance48.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.HeaderAppearance = appearance48;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.RowSelectorAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.SelectedRowAppearance = appearance50;
			this.uGridSalidaNoRegistrados.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridSalidaNoRegistrados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalidaNoRegistrados.Location = new System.Drawing.Point(712, 344);
			this.uGridSalidaNoRegistrados.Name = "uGridSalidaNoRegistrados";
			this.uGridSalidaNoRegistrados.Size = new System.Drawing.Size(346, 298);
			this.uGridSalidaNoRegistrados.TabIndex = 661;
			this.uGridSalidaNoRegistrados.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalidaNoRegistrados_CellChange);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(System.DateTime);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(bool);
			ultraDataColumn38.DefaultValue = false;
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38});
			// 
			// ultraGrid4
			// 
			this.ultraGrid4.Cursor = System.Windows.Forms.Cursors.Default;
			appearance51.BackColor = System.Drawing.Color.White;
			this.ultraGrid4.DisplayLayout.Appearance = appearance51;
			this.ultraGrid4.DisplayLayout.AutoFitColumns = true;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid4.DisplayLayout.Override.ActiveRowAppearance = appearance52;
			this.ultraGrid4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid4.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance54.FontData.BoldAsString = "True";
			appearance54.FontData.Name = "Arial";
			appearance54.FontData.SizeInPoints = 8.25F;
			appearance54.ForeColor = System.Drawing.Color.White;
			appearance54.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid4.DisplayLayout.Override.HeaderAppearance = appearance54;
			this.ultraGrid4.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance55.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance55.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.RowAlternateAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.RowSelectorAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.SelectedRowAppearance = appearance57;
			this.ultraGrid4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid4.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid4.Name = "ultraGrid4";
			this.ultraGrid4.Size = new System.Drawing.Size(346, 296);
			this.ultraGrid4.TabIndex = 659;
			this.ultraGrid4.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
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
			// btnConciliar
			// 
			this.btnConciliar.CausesValidation = false;
			this.btnConciliar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConciliar.Location = new System.Drawing.Point(152, 672);
			this.btnConciliar.Name = "btnConciliar";
			this.btnConciliar.TabIndex = 663;
			this.btnConciliar.Text = "&Conciliar";
			this.btnConciliar.Click += new System.EventHandler(this.btnConciliar_Click);
			// 
			// btnExportarEntradaSS
			// 
			this.btnExportarEntradaSS.CausesValidation = false;
			this.btnExportarEntradaSS.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarEntradaSS.Location = new System.Drawing.Point(360, 312);
			this.btnExportarEntradaSS.Name = "btnExportarEntradaSS";
			this.btnExportarEntradaSS.Size = new System.Drawing.Size(232, 23);
			this.btnExportarEntradaSS.TabIndex = 664;
			this.btnExportarEntradaSS.Text = "&Exportar Entrada Solo Sistema";
			this.btnExportarEntradaSS.Click += new System.EventHandler(this.btnExportarEntradaSS_Click);
			// 
			// btnExportarEntradaEC
			// 
			this.btnExportarEntradaEC.CausesValidation = false;
			this.btnExportarEntradaEC.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarEntradaEC.Location = new System.Drawing.Point(712, 312);
			this.btnExportarEntradaEC.Name = "btnExportarEntradaEC";
			this.btnExportarEntradaEC.Size = new System.Drawing.Size(232, 23);
			this.btnExportarEntradaEC.TabIndex = 665;
			this.btnExportarEntradaEC.Text = "&Exportar Entrada Estado De Cuenta";
			this.btnExportarEntradaEC.Click += new System.EventHandler(this.btnExportarEntradaEC_Click);
			// 
			// btnExportarSalidaSS
			// 
			this.btnExportarSalidaSS.CausesValidation = false;
			this.btnExportarSalidaSS.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarSalidaSS.Location = new System.Drawing.Point(360, 648);
			this.btnExportarSalidaSS.Name = "btnExportarSalidaSS";
			this.btnExportarSalidaSS.Size = new System.Drawing.Size(232, 23);
			this.btnExportarSalidaSS.TabIndex = 666;
			this.btnExportarSalidaSS.Text = "&Exportar Salida Solo Sistema";
			this.btnExportarSalidaSS.Click += new System.EventHandler(this.btnExportarSalidaSS_Click);
			// 
			// btnExportarSalidaEC
			// 
			this.btnExportarSalidaEC.CausesValidation = false;
			this.btnExportarSalidaEC.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarSalidaEC.Location = new System.Drawing.Point(712, 648);
			this.btnExportarSalidaEC.Name = "btnExportarSalidaEC";
			this.btnExportarSalidaEC.Size = new System.Drawing.Size(232, 23);
			this.btnExportarSalidaEC.TabIndex = 667;
			this.btnExportarSalidaEC.Text = "&Exportar Salida Estado De Cuenta";
			this.btnExportarSalidaEC.Click += new System.EventHandler(this.btnExportarSalidaEC_Click);
			// 
			// btnProcesarModificacion
			// 
			this.btnProcesarModificacion.CausesValidation = false;
			this.btnProcesarModificacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesarModificacion.Location = new System.Drawing.Point(8, 672);
			this.btnProcesarModificacion.Name = "btnProcesarModificacion";
			this.btnProcesarModificacion.Size = new System.Drawing.Size(136, 23);
			this.btnProcesarModificacion.TabIndex = 668;
			this.btnProcesarModificacion.Text = "&Procesar Modificacion";
			this.btnProcesarModificacion.Click += new System.EventHandler(this.btnProcesarModificacion_Click);
			// 
			// frmComparaEstadosDeCuenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1064, 704);
			this.Controls.Add(this.btnProcesarModificacion);
			this.Controls.Add(this.btnExportarSalidaEC);
			this.Controls.Add(this.btnExportarSalidaSS);
			this.Controls.Add(this.btnExportarEntradaEC);
			this.Controls.Add(this.btnExportarEntradaSS);
			this.Controls.Add(this.btnConciliar);
			this.Controls.Add(this.uGridEntradaNoRegistrados);
			this.Controls.Add(this.uGridSalidaNoRegistrados);
			this.Controls.Add(this.uGridEntradaNoConciliados);
			this.Controls.Add(this.uGridSalidasNoConciliados);
			this.Controls.Add(this.uGridSalidaConciliados);
			this.Controls.Add(this.uGridEntradaConciliados);
			this.Controls.Add(this.ultraGrid4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmComparaEstadosDeCuenta";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Compara Movimientos Bancarios con Estado de Cuenta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmComparaEstadosDeCuenta_KeyDown);
			this.Load += new System.EventHandler(this.frmComparaEstadosDeCuenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaNoConciliados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidasNoConciliados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidaConciliados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaConciliados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradaNoRegistrados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidaNoRegistrados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void frmComparaEstadosDeCuenta_Load(object sender, System.EventArgs e)
		{	
			#region Entradas
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEntradas in uGridEntradas.Rows.All)
			{
				string sENumero = drEntradas.Cells["Cheque"].Value.ToString();
				decimal dEValor = Convert.ToDecimal(drEntradas.Cells["Valor"].Value);

				if ((bool)drEntradas.Cells["Confirmado"].Value)
				{
					this.uGridEntradaConciliados.Rows.Band.AddNew();
					this.uGridEntradaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drEntradas.Cells["idAsiento"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drEntradas.Cells["idDetAsiento"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drEntradas.Cells["Fecha"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["Movimiento"].Value = drEntradas.Cells["Movimiento"].Value.ToString();
					this.uGridEntradaConciliados.ActiveRow.Cells["Cheque"].Value = drEntradas.Cells["Cheque"].Value.ToString();
					this.uGridEntradaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drEntradas.Cells["Valor"].Value);
				}
				else
				{
					bool bConciliado = false;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drECEntradas in uGridECEntradas.Rows.All)
					{
						string sEECNumero = drECEntradas.Cells["Numero"].Value.ToString();
						decimal dEECValor = Convert.ToDecimal(drECEntradas.Cells["Valor"].Value);
										
						if (sEECNumero == sENumero && Convert.ToDecimal(dEECValor) == Convert.ToDecimal(dEValor))
						{
							this.uGridEntradaConciliados.Rows.Band.AddNew();
							this.uGridEntradaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drEntradas.Cells["idAsiento"].Value;
							this.uGridEntradaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drEntradas.Cells["idDetAsiento"].Value;
							this.uGridEntradaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drEntradas.Cells["Fecha"].Value;
							this.uGridEntradaConciliados.ActiveRow.Cells["Movimiento"].Value = drEntradas.Cells["Movimiento"].Value.ToString();
							this.uGridEntradaConciliados.ActiveRow.Cells["Cheque"].Value = drEntradas.Cells["Cheque"].Value.ToString();
							this.uGridEntradaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drEntradas.Cells["Valor"].Value);					

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 1 Where idEstadoDeCuentaBancaria = {0}", (int)drECEntradas.Cells["idEstadoDeCuentaBancaria"].Value));

							bConciliado = true;

							break;
						}
					}

					if (!bConciliado)
					{
						this.uGridEntradaNoConciliados.Rows.Band.AddNew();
						this.uGridEntradaNoConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drEntradas.Cells["idAsiento"].Value;
						this.uGridEntradaNoConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drEntradas.Cells["idDetAsiento"].Value;
						this.uGridEntradaNoConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drEntradas.Cells["Fecha"].Value;
						this.uGridEntradaNoConciliados.ActiveRow.Cells["Movimiento"].Value = drEntradas.Cells["Movimiento"].Value.ToString();
						this.uGridEntradaNoConciliados.ActiveRow.Cells["Cheque"].Value = drEntradas.Cells["Cheque"].Value.ToString();
						this.uGridEntradaNoConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drEntradas.Cells["Valor"].Value);											
					}
				}
			}
			#endregion Entradas
			
			#region Salidas
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSalidas in uGridSalidas.Rows.All)
			{
				string sSNumero = drSalidas.Cells["Cheque"].Value.ToString();
				decimal dSValor = Convert.ToDecimal(drSalidas.Cells["Valor"].Value);

				if ((bool)drSalidas.Cells["Confirmado"].Value)
				{
					this.uGridSalidaConciliados.Rows.Band.AddNew();
					this.uGridSalidaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drSalidas.Cells["idAsiento"].Value;;
					this.uGridSalidaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drSalidas.Cells["idDetAsiento"].Value;;
					this.uGridSalidaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drSalidas.Cells["Fecha"].Value;
					this.uGridSalidaConciliados.ActiveRow.Cells["Movimiento"].Value = drSalidas.Cells["Movimiento"].Value.ToString();
					this.uGridSalidaConciliados.ActiveRow.Cells["Cheque"].Value = drSalidas.Cells["Cheque"].Value.ToString();
					this.uGridSalidaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drSalidas.Cells["Valor"].Value);
				}
				else
				{
					bool bConciliado = false;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drECSalidas in uGridECSalidas.Rows.All)
					{
						string sSECNumero = drECSalidas.Cells["Numero"].Value.ToString();
						decimal dSECValor = Convert.ToDecimal(drECSalidas.Cells["Valor"].Value);
										
						if (sSECNumero == sSNumero && Convert.ToDecimal(dSECValor) == Convert.ToDecimal(dSValor))
						{
							this.uGridSalidaConciliados.Rows.Band.AddNew();
							this.uGridSalidaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drSalidas.Cells["idAsiento"].Value;;
							this.uGridSalidaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drSalidas.Cells["idDetAsiento"].Value;;
							this.uGridSalidaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drSalidas.Cells["Fecha"].Value;
							this.uGridSalidaConciliados.ActiveRow.Cells["Movimiento"].Value = drSalidas.Cells["Movimiento"].Value.ToString();
							this.uGridSalidaConciliados.ActiveRow.Cells["Cheque"].Value = drSalidas.Cells["Cheque"].Value.ToString();
							this.uGridSalidaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drSalidas.Cells["Valor"].Value);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 1 Where idEstadoDeCuentaBancaria = {0}", (int)drECSalidas.Cells["idEstadoDeCuentaBancaria"].Value));

							bConciliado = true;

							break;
						}
					}

					if (!bConciliado)
					{
						this.uGridSalidasNoConciliados.Rows.Band.AddNew();
						this.uGridSalidasNoConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drSalidas.Cells["idAsiento"].Value;;
						this.uGridSalidasNoConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drSalidas.Cells["idDetAsiento"].Value;;
						this.uGridSalidasNoConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drSalidas.Cells["Fecha"].Value;
						this.uGridSalidasNoConciliados.ActiveRow.Cells["Movimiento"].Value = drSalidas.Cells["Movimiento"].Value.ToString();
						this.uGridSalidasNoConciliados.ActiveRow.Cells["Cheque"].Value = drSalidas.Cells["Cheque"].Value.ToString();
						this.uGridSalidasNoConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drSalidas.Cells["Valor"].Value);
					}
				}
			}			
			#endregion Salidas

			#region Entradas En Estado de cuenta que no estan en el sistema
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drECNREntradas in uGridECEntradas.Rows.All)
			{
				string sEECNRNumero = drECNREntradas.Cells["Numero"].Value.ToString();
				decimal dEECNRValor = Convert.ToDecimal(drECNREntradas.Cells["Valor"].Value);

				bool bExiste = false;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNREntradas in uGridEntradas.Rows.All)
				{
					if (!(bool)drNREntradas.Cells["Confirmado"].Value)
					{
						string sENRNumero = drNREntradas.Cells["Cheque"].Value.ToString();
						decimal dENRValor = Convert.ToDecimal(drNREntradas.Cells["Valor"].Value);
										
						if (sENRNumero == sEECNRNumero && Convert.ToDecimal(dENRValor) == Convert.ToDecimal(dEECNRValor))
						{
							bExiste = true;
							break;
						}						
					}					
				}	
			
				if (!bExiste)
				{
					if (!(bool)drECNREntradas.Cells["Sel"].Value)
					{
						this.uGridEntradaNoRegistrados.Rows.Band.AddNew();
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["idEstadoDeCuentaBancaria"].Value = (int)drECNREntradas.Cells["idEstadoDeCuentaBancaria"].Value;
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["Fecha"].Value = (DateTime)drECNREntradas.Cells["Fecha"].Value;
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["Tipo"].Value = drECNREntradas.Cells["Tipo"].Value.ToString();
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["Numero"].Value = drECNREntradas.Cells["Numero"].Value.ToString();
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drECNREntradas.Cells["Valor"].Value);
						this.uGridEntradaNoRegistrados.ActiveRow.Cells["Sel"].Value = (bool)drECNREntradas.Cells["Sel"].Value;

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 2 Where idEstadoDeCuentaBancaria = {0}", (int)drECNREntradas.Cells["idEstadoDeCuentaBancaria"].Value));
					}
				}
			}
			#endregion Entradas En Estado de cuenta que no estan en el sistema

			#region Salidas En Estado de cuenta que no estan en el sistema
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSCNRSalidas in uGridECSalidas.Rows.All)
			{
				string sSECNRNumero = drSCNRSalidas.Cells["Numero"].Value.ToString();
				decimal dSECNRValor = Convert.ToDecimal(drSCNRSalidas.Cells["Valor"].Value);

				bool bExiste = false;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNRSalidas in uGridSalidas.Rows.All)
				{
					string sSNRNumero = drNRSalidas.Cells["Cheque"].Value.ToString();
					decimal dSNRValor = Convert.ToDecimal(drNRSalidas.Cells["Valor"].Value);
										
					if (sSNRNumero == sSECNRNumero && Convert.ToDecimal(dSNRValor) == Convert.ToDecimal(dSECNRValor))
					{
						bExiste = true;

						break;
					}
				}
			
				if (!bExiste)
				{
					if (!(bool)drSCNRSalidas.Cells["Sel"].Value)
					{
						this.uGridSalidaNoRegistrados.Rows.Band.AddNew();
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["idEstadoDeCuentaBancaria"].Value = (int)drSCNRSalidas.Cells["idEstadoDeCuentaBancaria"].Value;
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["Fecha"].Value = (DateTime)drSCNRSalidas.Cells["Fecha"].Value;
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["Tipo"].Value = drSCNRSalidas.Cells["Tipo"].Value.ToString();
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["Numero"].Value = drSCNRSalidas.Cells["Numero"].Value.ToString();
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drSCNRSalidas.Cells["Valor"].Value);
						this.uGridSalidaNoRegistrados.ActiveRow.Cells["Sel"].Value = (bool)drSCNRSalidas.Cells["Sel"].Value;

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 2 Where idEstadoDeCuentaBancaria = {0}", (int)drSCNRSalidas.Cells["idEstadoDeCuentaBancaria"].Value));
					}
				}
			}
			#endregion Salidas En Estado de cuenta que no estan en el sistema
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmComparaEstadosDeCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnExportarEntradaSS_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridEntradaNoConciliados, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnExportarEntradaEC_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridEntradaNoRegistrados, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnExportarSalidaSS_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridSalidasNoConciliados, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnExportarSalidaEC_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridSalidaNoRegistrados, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void uGridEntradaNoRegistrados_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idEstadoDeCuentaBancaria"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridEntradaNoRegistrados.UpdateData();

				string sSQL = string.Format("Exec CBActualizaEstadoDeCuentas {0}, {1}", (int)e.Cell.Row.Cells["idEstadoDeCuentaBancaria"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void uGridSalidaNoRegistrados_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idEstadoDeCuentaBancaria"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridSalidaNoRegistrados.UpdateData();

				string sSQL = string.Format("Exec CBActualizaEstadoDeCuentas {0}, {1}", (int)e.Cell.Row.Cells["idEstadoDeCuentaBancaria"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void uGridEntradaConciliados_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Cheque")
			{
				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec CBActualizaNumeroChequeDeposito '{0}', {1}", e.Cell.Row.Cells["Cheque"].Value.ToString(), (int)e.Cell.Row.Cells["idAsiento"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void uGridSalidaConciliados_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Cheque")
			{
				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec CBActualizaNumeroChequeDeposito '{0}', {1}", e.Cell.Row.Cells["Cheque"].Value.ToString(), (int)e.Cell.Row.Cells["idAsiento"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void uGridEntradaNoConciliados_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Cheque")
			{
				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec CBActualizaNumeroChequeDeposito '{0}', {1}", e.Cell.Row.Cells["Cheque"].Value.ToString(), (int)e.Cell.Row.Cells["idAsiento"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void uGridSalidasNoConciliados_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Cheque")
			{
				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec CBActualizaNumeroChequeDeposito '{0}', {1}", e.Cell.Row.Cells["Cheque"].Value.ToString(), (int)e.Cell.Row.Cells["idAsiento"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void btnConciliar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btnProcesarModificacion_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drENC in this.uGridEntradaNoConciliados.Rows.All)
			{
				if ((bool)drENC.Cells["Confirmado"].Value)
				{
					this.uGridEntradaConciliados.Rows.Band.AddNew();
					this.uGridEntradaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drENC.Cells["idAsiento"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drENC.Cells["idDetAsiento"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drENC.Cells["Fecha"].Value;
					this.uGridEntradaConciliados.ActiveRow.Cells["Movimiento"].Value = drENC.Cells["Movimiento"].Value.ToString();
					this.uGridEntradaConciliados.ActiveRow.Cells["Cheque"].Value = drENC.Cells["Cheque"].Value.ToString();
					this.uGridEntradaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drENC.Cells["Valor"].Value);			

					drENC.Delete(false);
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSNC in this.uGridSalidasNoConciliados.Rows.All)
			{
				if ((bool)drSNC.Cells["Confirmado"].Value)
				{
					this.uGridSalidaConciliados.Rows.Band.AddNew();
					this.uGridSalidaConciliados.ActiveRow.Cells["idAsiento"].Value = (int)drSNC.Cells["idAsiento"].Value;
					this.uGridSalidaConciliados.ActiveRow.Cells["idDetAsiento"].Value = (int)drSNC.Cells["idDetAsiento"].Value;
					this.uGridSalidaConciliados.ActiveRow.Cells["Fecha"].Value = (DateTime)drSNC.Cells["Fecha"].Value;
					this.uGridSalidaConciliados.ActiveRow.Cells["Movimiento"].Value = drSNC.Cells["Movimiento"].Value.ToString();
					this.uGridSalidaConciliados.ActiveRow.Cells["Cheque"].Value = drSNC.Cells["Cheque"].Value.ToString();
					this.uGridSalidaConciliados.ActiveRow.Cells["Valor"].Value = Convert.ToDecimal(drSNC.Cells["Valor"].Value);			

					drSNC.Delete(false);
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drENR in this.uGridEntradaNoRegistrados.Rows.All)
			{
				if ((bool)drENR.Cells["Sel"].Value) 
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 1 Where idEstadoDeCuentaBancaria = {0}", (int)drENR.Cells["idEstadoDeCuentaBancaria"].Value));
					
					drENR.Delete(false);
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSNR in this.uGridSalidaNoRegistrados.Rows.All)
			{
				if ((bool)drSNR.Cells["Sel"].Value) 
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update EstadosDeCuentasBancarias Set Estado = 1 Where idEstadoDeCuentaBancaria = {0}", (int)drSNR.Cells["idEstadoDeCuentaBancaria"].Value));
				
					drSNR.Delete(false);
				}
			}
		}
	}
}

