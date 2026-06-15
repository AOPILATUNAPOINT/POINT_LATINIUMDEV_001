using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReportesEstadisticosCartera.
	/// </summary>
	public class frmCre_ReportesEstadisticosCartera : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCartera;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoVencido;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRatioVencido;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRatio30;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo30;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoFDG;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCoberturaFDG;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVariaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGtEvolucionRiesgo;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridPrimerasCuotas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridResumenColocaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridAnalisisdeCosechas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRetiros;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridFDG;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReportesEstadisticosCartera()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operaciones");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresoFDG");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fdg_actual");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OperacionesCaenFDG");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EgresosFDG");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VariacionFDG");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IngresoFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fdg_actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OperacionesCaenFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EgresosFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VariacionFDG");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoPeriodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RiesgoTotal");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoVencido");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ratio");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RiesgoTotal");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoVencido");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ratio");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RiesgoTotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RiesgoTotal", 1, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoVencido", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoVencido", 2, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Ratio", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Ratio", 3, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RiesgoTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ratio");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresoFDG");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EgresoFDG");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PIngresoFDG");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PEgresoFDG");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PEgresoProy");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorEgresarFDG");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCosecha");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 1, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IngresoFDG", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IngresoFDG", 2, true);
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "EgresoFDG", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "EgresoFDG", 3, true);
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoCosecha", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoCosecha", 8, true);
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IngresoFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EgresoFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PIngresoFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PEgresoFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PEgresoProy");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorEgresarFDG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCosecha");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clientes");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 1, true);
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Clientes", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Clientes", 2, true);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Promedio", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Promedio", 3, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Clientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promedio");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoMorosidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NoDevengado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VencidoTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RiesgoTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ParticipacionSaldo");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReportesEstadisticosCartera));
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.utpEstado = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtCoberturaFDG = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtSaldoFDG = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtSaldoVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRatioVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRatio30 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldo30 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridCartera = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridVariaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label15 = new System.Windows.Forms.Label();
			this.uGridPrimerasCuotas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGtEvolucionRiesgo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridAnalisisdeCosechas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridResumenColocaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.uGridFDG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridRetiros = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btnVer = new System.Windows.Forms.Button();
			this.utpEstado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCoberturaFDG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFDG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatioVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatio30)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo30)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrimerasCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGtEvolucionRiesgo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridAnalisisdeCosechas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumenColocaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridFDG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetiros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			this.SuspendLayout();
			// 
			// utpEstado
			// 
			this.utpEstado.Controls.Add(this.label12);
			this.utpEstado.Controls.Add(this.label11);
			this.utpEstado.Controls.Add(this.label10);
			this.utpEstado.Controls.Add(this.label9);
			this.utpEstado.Controls.Add(this.txtCoberturaFDG);
			this.utpEstado.Controls.Add(this.groupBox1);
			this.utpEstado.Controls.Add(this.label6);
			this.utpEstado.Controls.Add(this.label5);
			this.utpEstado.Controls.Add(this.txtSaldoFDG);
			this.utpEstado.Controls.Add(this.label2);
			this.utpEstado.Controls.Add(this.label1);
			this.utpEstado.Controls.Add(this.label8);
			this.utpEstado.Controls.Add(this.label3);
			this.utpEstado.Controls.Add(this.txtSaldoVencido);
			this.utpEstado.Controls.Add(this.txtRatioVencido);
			this.utpEstado.Controls.Add(this.txtRatio30);
			this.utpEstado.Controls.Add(this.txtSaldo30);
			this.utpEstado.Controls.Add(this.uGridCartera);
			this.utpEstado.Controls.Add(this.uGridVariaciones);
			this.utpEstado.Location = new System.Drawing.Point(1, 20);
			this.utpEstado.Name = "utpEstado";
			this.utpEstado.Size = new System.Drawing.Size(1110, 403);
			this.utpEstado.Paint += new System.Windows.Forms.PaintEventHandler(this.utpEstado_Paint);
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(1024, 74);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(14, 16);
			this.label12.TabIndex = 758;
			this.label12.Text = "%";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(1024, 122);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(14, 16);
			this.label11.TabIndex = 757;
			this.label11.Text = "%";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(1024, 170);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(14, 16);
			this.label10.TabIndex = 756;
			this.label10.Text = "%";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(776, 170);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(55, 16);
			this.label9.TabIndex = 754;
			this.label9.Text = "Cobertura";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCoberturaFDG
			// 
			this.txtCoberturaFDG.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCoberturaFDG.Appearance = appearance1;
			this.txtCoberturaFDG.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCoberturaFDG.Enabled = false;
			this.txtCoberturaFDG.FormatString = "#,##0.00";
			this.txtCoberturaFDG.Location = new System.Drawing.Point(912, 168);
			this.txtCoberturaFDG.Name = "txtCoberturaFDG";
			this.txtCoberturaFDG.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCoberturaFDG.PromptChar = ' ';
			this.txtCoberturaFDG.Size = new System.Drawing.Size(100, 21);
			this.txtCoberturaFDG.TabIndex = 755;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(768, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(272, 8);
			this.groupBox1.TabIndex = 753;
			this.groupBox1.TabStop = false;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(768, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(272, 16);
			this.label6.TabIndex = 752;
			this.label6.Text = "SALDOS Y RATIOS";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(776, 146);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(60, 16);
			this.label5.TabIndex = 750;
			this.label5.Text = "Saldo FDG";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoFDG
			// 
			this.txtSaldoFDG.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoFDG.Appearance = appearance2;
			this.txtSaldoFDG.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoFDG.Enabled = false;
			this.txtSaldoFDG.FormatString = "#,##0.00";
			this.txtSaldoFDG.Location = new System.Drawing.Point(912, 144);
			this.txtSaldoFDG.Name = "txtSaldoFDG";
			this.txtSaldoFDG.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoFDG.PromptChar = ' ';
			this.txtSaldoFDG.Size = new System.Drawing.Size(100, 21);
			this.txtSaldoFDG.TabIndex = 751;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(776, 50);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(77, 16);
			this.label2.TabIndex = 748;
			this.label2.Text = "Saldo Vencido";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(776, 74);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(75, 16);
			this.label1.TabIndex = 746;
			this.label1.Text = "Ratio Vencido";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(776, 122);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(56, 16);
			this.label8.TabIndex = 744;
			this.label8.Text = "Ratio 30 +";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(776, 98);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(58, 16);
			this.label3.TabIndex = 742;
			this.label3.Text = "Saldo 30 +";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoVencido
			// 
			this.txtSaldoVencido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoVencido.Appearance = appearance3;
			this.txtSaldoVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoVencido.Enabled = false;
			this.txtSaldoVencido.FormatString = "#,##0.00";
			this.txtSaldoVencido.Location = new System.Drawing.Point(912, 48);
			this.txtSaldoVencido.Name = "txtSaldoVencido";
			this.txtSaldoVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoVencido.PromptChar = ' ';
			this.txtSaldoVencido.Size = new System.Drawing.Size(100, 21);
			this.txtSaldoVencido.TabIndex = 749;
			// 
			// txtRatioVencido
			// 
			this.txtRatioVencido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRatioVencido.Appearance = appearance4;
			this.txtRatioVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRatioVencido.Enabled = false;
			this.txtRatioVencido.FormatString = "#,##0.00";
			this.txtRatioVencido.Location = new System.Drawing.Point(912, 72);
			this.txtRatioVencido.Name = "txtRatioVencido";
			this.txtRatioVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRatioVencido.PromptChar = ' ';
			this.txtRatioVencido.Size = new System.Drawing.Size(100, 21);
			this.txtRatioVencido.TabIndex = 747;
			// 
			// txtRatio30
			// 
			this.txtRatio30.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRatio30.Appearance = appearance5;
			this.txtRatio30.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRatio30.Enabled = false;
			this.txtRatio30.FormatString = "#,##0.00";
			this.txtRatio30.Location = new System.Drawing.Point(912, 120);
			this.txtRatio30.Name = "txtRatio30";
			this.txtRatio30.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRatio30.PromptChar = ' ';
			this.txtRatio30.Size = new System.Drawing.Size(100, 21);
			this.txtRatio30.TabIndex = 745;
			// 
			// txtSaldo30
			// 
			this.txtSaldo30.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo30.Appearance = appearance6;
			this.txtSaldo30.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo30.Enabled = false;
			this.txtSaldo30.FormatString = "#,##0.00";
			this.txtSaldo30.Location = new System.Drawing.Point(912, 96);
			this.txtSaldo30.Name = "txtSaldo30";
			this.txtSaldo30.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo30.PromptChar = ' ';
			this.txtSaldo30.Size = new System.Drawing.Size(100, 21);
			this.txtSaldo30.TabIndex = 743;
			// 
			// uGridCartera
			// 
			this.uGridCartera.Cursor = System.Windows.Forms.Cursors.Default;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridCartera.DisplayLayout.Appearance = appearance7;
			this.uGridCartera.DisplayLayout.AutoFitColumns = true;
			this.uGridCartera.DisplayLayout.GroupByBox.Hidden = true;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCartera.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridCartera.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCartera.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridCartera.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCartera.DisplayLayout.Override.CellAppearance = appearance10;
			this.uGridCartera.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCartera.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridCartera.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCartera.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridCartera.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridCartera.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCartera.Location = new System.Drawing.Point(8, 8);
			this.uGridCartera.Name = "uGridCartera";
			this.uGridCartera.Size = new System.Drawing.Size(741, 200);
			this.uGridCartera.TabIndex = 741;
			this.uGridCartera.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCartera_ClickCellButton);
			// 
			// uGridVariaciones
			// 
			this.uGridVariaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVariaciones.DataSource = this.ultraDataSource4;
			appearance15.BackColor = System.Drawing.Color.White;
			this.uGridVariaciones.DisplayLayout.Appearance = appearance15;
			this.uGridVariaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn1.CellAppearance = appearance16;
			ultraGridColumn1.Format = "#,##0.00";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 99;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance17;
			ultraGridColumn2.Format = "#,##0";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 102;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance18;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.Caption = "Ingreso F D G";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 99;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance19;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "% F D G Actual";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 99;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance20;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.Caption = "Op En F D G";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 99;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance21;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Egresos F D G";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance22;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Variacion F D G";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 102;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "Variaciones";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridVariaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVariaciones.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.uGridVariaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVariaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVariaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridVariaciones.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVariaciones.DisplayLayout.Override.CellAppearance = appearance25;
			this.uGridVariaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8.5F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVariaciones.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.uGridVariaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariaciones.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariaciones.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.uGridVariaciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridVariaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVariaciones.Location = new System.Drawing.Point(8, 216);
			this.uGridVariaciones.Name = "uGridVariaciones";
			this.uGridVariaciones.Size = new System.Drawing.Size(721, 120);
			this.uGridVariaciones.TabIndex = 759;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(System.Decimal);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.label15);
			this.pgEntrega.Controls.Add(this.uGridPrimerasCuotas);
			this.pgEntrega.Controls.Add(this.uGtEvolucionRiesgo);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(1110, 403);
			// 
			// label15
			// 
			this.label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label15.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(400, 8);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(704, 17);
			this.label15.TabIndex = 762;
			this.label15.Text = "PRIMERAS CUOTAS";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridPrimerasCuotas
			// 
			this.uGridPrimerasCuotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPrimerasCuotas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance29.BackColor = System.Drawing.Color.White;
			this.uGridPrimerasCuotas.DisplayLayout.Appearance = appearance29;
			ultraGridBand2.AddButtonCaption = "Serie";
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Periodo";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance30;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Riesgo Total";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance31;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Saldo Vencido";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance32;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridPrimerasCuotas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPrimerasCuotas.DisplayLayout.Override.ActiveRowAppearance = appearance33;
			this.uGridPrimerasCuotas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPrimerasCuotas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPrimerasCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance34.BackColor = System.Drawing.Color.Transparent;
			this.uGridPrimerasCuotas.DisplayLayout.Override.CardAreaAppearance = appearance34;
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPrimerasCuotas.DisplayLayout.Override.CellAppearance = appearance35;
			this.uGridPrimerasCuotas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance36.FontData.BoldAsString = "True";
			appearance36.FontData.Name = "Arial";
			appearance36.FontData.SizeInPoints = 8.5F;
			appearance36.ForeColor = System.Drawing.Color.White;
			appearance36.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPrimerasCuotas.DisplayLayout.Override.HeaderAppearance = appearance36;
			this.uGridPrimerasCuotas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrimerasCuotas.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrimerasCuotas.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.uGridPrimerasCuotas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridPrimerasCuotas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPrimerasCuotas.Location = new System.Drawing.Point(400, 32);
			this.uGridPrimerasCuotas.Name = "uGridPrimerasCuotas";
			this.uGridPrimerasCuotas.Size = new System.Drawing.Size(704, 368);
			this.uGridPrimerasCuotas.TabIndex = 758;
			// 
			// uGtEvolucionRiesgo
			// 
			this.uGtEvolucionRiesgo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGtEvolucionRiesgo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGtEvolucionRiesgo.DataSource = this.ultraDataSource5;
			appearance39.BackColor = System.Drawing.Color.White;
			this.uGtEvolucionRiesgo.DisplayLayout.Appearance = appearance39;
			this.uGtEvolucionRiesgo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Serie";
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance40;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 61;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance41;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Riesgo Total";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance42;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Saldo Vencido";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance43;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 100;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand3.Header.Caption = "Evolución Riesgo Total y Ratio de Vencida";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance44;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance45;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance46;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGtEvolucionRiesgo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance47.ForeColor = System.Drawing.Color.Black;
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.ActiveRowAppearance = appearance47;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance48.BackColor = System.Drawing.Color.Transparent;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.CardAreaAppearance = appearance48;
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.CellAppearance = appearance49;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 8.5F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.HeaderAppearance = appearance50;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.RowSelectorAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.SelectedRowAppearance = appearance52;
			this.uGtEvolucionRiesgo.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGtEvolucionRiesgo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGtEvolucionRiesgo.Location = new System.Drawing.Point(8, 8);
			this.uGtEvolucionRiesgo.Name = "uGtEvolucionRiesgo";
			this.uGtEvolucionRiesgo.Size = new System.Drawing.Size(381, 392);
			this.uGtEvolucionRiesgo.TabIndex = 757;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.uGridAnalisisdeCosechas);
			this.ultraTabPageControl1.Controls.Add(this.uGridResumenColocaciones);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1110, 403);
			// 
			// uGridAnalisisdeCosechas
			// 
			this.uGridAnalisisdeCosechas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAnalisisdeCosechas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAnalisisdeCosechas.DataSource = this.ultraDataSource3;
			appearance53.BackColor = System.Drawing.Color.White;
			this.uGridAnalisisdeCosechas.DisplayLayout.Appearance = appearance53;
			this.uGridAnalisisdeCosechas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "Serie";
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance54;
			ultraGridColumn16.Header.Caption = "Cosecha";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Width = 71;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance55;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance56;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Ingreso F D G";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 93;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance57;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Egreso F D G";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 93;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance58;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "% Ingreso F D G";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 93;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance59;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "% Egreso F D G";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 93;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance60;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 93;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance61;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 93;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance62;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Saldo Cosecha";
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Width = 93;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			ultraGridBand4.Header.Caption = "Análisis de Cosechas";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance63;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance64;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance65;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance66;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridAnalisisdeCosechas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.ActiveRowAppearance = appearance67;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance68.BackColor = System.Drawing.Color.Transparent;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.CardAreaAppearance = appearance68;
			appearance69.ForeColor = System.Drawing.Color.Black;
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.CellAppearance = appearance69;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance70.FontData.BoldAsString = "True";
			appearance70.FontData.Name = "Arial";
			appearance70.FontData.SizeInPoints = 8.5F;
			appearance70.ForeColor = System.Drawing.Color.White;
			appearance70.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.HeaderAppearance = appearance70;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.RowSelectorAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.SelectedRowAppearance = appearance72;
			this.uGridAnalisisdeCosechas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridAnalisisdeCosechas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAnalisisdeCosechas.Location = new System.Drawing.Point(440, 8);
			this.uGridAnalisisdeCosechas.Name = "uGridAnalisisdeCosechas";
			this.uGridAnalisisdeCosechas.Size = new System.Drawing.Size(656, 392);
			this.uGridAnalisisdeCosechas.TabIndex = 760;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// uGridResumenColocaciones
			// 
			this.uGridResumenColocaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridResumenColocaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridResumenColocaciones.DataSource = this.ultraDataSource1;
			appearance73.BackColor = System.Drawing.Color.White;
			this.uGridResumenColocaciones.DisplayLayout.Appearance = appearance73;
			this.uGridResumenColocaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "Serie";
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance74;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Width = 101;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance75;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance76;
			ultraGridColumn27.Format = "#,##0";
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance77;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Width = 100;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			ultraGridBand5.Header.Caption = "Resumen Colocaciones";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance78;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance79;
			summarySettings9.DisplayFormat = "{0: #,##0}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance80;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridResumenColocaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance81.ForeColor = System.Drawing.Color.Black;
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumenColocaciones.DisplayLayout.Override.ActiveRowAppearance = appearance81;
			this.uGridResumenColocaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridResumenColocaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridResumenColocaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance82.BackColor = System.Drawing.Color.Transparent;
			this.uGridResumenColocaciones.DisplayLayout.Override.CardAreaAppearance = appearance82;
			appearance83.ForeColor = System.Drawing.Color.Black;
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumenColocaciones.DisplayLayout.Override.CellAppearance = appearance83;
			this.uGridResumenColocaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance84.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance84.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance84.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance84.FontData.BoldAsString = "True";
			appearance84.FontData.Name = "Arial";
			appearance84.FontData.SizeInPoints = 8.5F;
			appearance84.ForeColor = System.Drawing.Color.White;
			appearance84.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResumenColocaciones.DisplayLayout.Override.HeaderAppearance = appearance84;
			this.uGridResumenColocaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance85.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance85.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumenColocaciones.DisplayLayout.Override.RowSelectorAppearance = appearance85;
			appearance86.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance86.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumenColocaciones.DisplayLayout.Override.SelectedRowAppearance = appearance86;
			this.uGridResumenColocaciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridResumenColocaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResumenColocaciones.Location = new System.Drawing.Point(8, 8);
			this.uGridResumenColocaciones.Name = "uGridResumenColocaciones";
			this.uGridResumenColocaciones.Size = new System.Drawing.Size(421, 392);
			this.uGridResumenColocaciones.TabIndex = 759;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.label14);
			this.ultraTabPageControl2.Controls.Add(this.label13);
			this.ultraTabPageControl2.Controls.Add(this.uGridFDG);
			this.ultraTabPageControl2.Controls.Add(this.uGridRetiros);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1110, 403);
			// 
			// label14
			// 
			this.label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label14.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(560, 8);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(544, 17);
			this.label14.TabIndex = 762;
			this.label14.Text = "CRÉDITOS EN FONDO DE GARANTÍA";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label13
			// 
			this.label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label13.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(8, 8);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(544, 17);
			this.label13.TabIndex = 761;
			this.label13.Text = "CRÉDITOS DE RETIRO";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridFDG
			// 
			this.uGridFDG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFDG.Cursor = System.Windows.Forms.Cursors.Default;
			appearance87.BackColor = System.Drawing.Color.White;
			this.uGridFDG.DisplayLayout.Appearance = appearance87;
			appearance88.ForeColor = System.Drawing.Color.Black;
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFDG.DisplayLayout.Override.ActiveRowAppearance = appearance88;
			this.uGridFDG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridFDG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridFDG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance89.BackColor = System.Drawing.Color.Transparent;
			this.uGridFDG.DisplayLayout.Override.CardAreaAppearance = appearance89;
			appearance90.ForeColor = System.Drawing.Color.Black;
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFDG.DisplayLayout.Override.CellAppearance = appearance90;
			this.uGridFDG.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance91.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance91.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance91.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance91.FontData.BoldAsString = "True";
			appearance91.FontData.Name = "Arial";
			appearance91.FontData.SizeInPoints = 8.5F;
			appearance91.ForeColor = System.Drawing.Color.White;
			appearance91.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFDG.DisplayLayout.Override.HeaderAppearance = appearance91;
			this.uGridFDG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance92.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance92.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFDG.DisplayLayout.Override.RowSelectorAppearance = appearance92;
			appearance93.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance93.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance93.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFDG.DisplayLayout.Override.SelectedRowAppearance = appearance93;
			this.uGridFDG.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridFDG.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFDG.Location = new System.Drawing.Point(560, 32);
			this.uGridFDG.Name = "uGridFDG";
			this.uGridFDG.Size = new System.Drawing.Size(544, 368);
			this.uGridFDG.TabIndex = 760;
			// 
			// uGridRetiros
			// 
			this.uGridRetiros.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRetiros.Cursor = System.Windows.Forms.Cursors.Default;
			appearance94.BackColor = System.Drawing.Color.White;
			this.uGridRetiros.DisplayLayout.Appearance = appearance94;
			appearance95.ForeColor = System.Drawing.Color.Black;
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRetiros.DisplayLayout.Override.ActiveRowAppearance = appearance95;
			this.uGridRetiros.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRetiros.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRetiros.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance96.BackColor = System.Drawing.Color.Transparent;
			this.uGridRetiros.DisplayLayout.Override.CardAreaAppearance = appearance96;
			appearance97.ForeColor = System.Drawing.Color.Black;
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRetiros.DisplayLayout.Override.CellAppearance = appearance97;
			this.uGridRetiros.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance98.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance98.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance98.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance98.FontData.BoldAsString = "True";
			appearance98.FontData.Name = "Arial";
			appearance98.FontData.SizeInPoints = 8.5F;
			appearance98.ForeColor = System.Drawing.Color.White;
			appearance98.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRetiros.DisplayLayout.Override.HeaderAppearance = appearance98;
			this.uGridRetiros.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance99.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance99.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetiros.DisplayLayout.Override.RowSelectorAppearance = appearance99;
			appearance100.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance100.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetiros.DisplayLayout.Override.SelectedRowAppearance = appearance100;
			this.uGridRetiros.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridRetiros.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRetiros.Location = new System.Drawing.Point(8, 32);
			this.uGridRetiros.Name = "uGridRetiros";
			this.uGridRetiros.Size = new System.Drawing.Size(544, 368);
			this.uGridRetiros.TabIndex = 759;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 744;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance101;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 745;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnImprimir
			// 
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1048, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 750;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 42);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 17);
			this.label7.TabIndex = 751;
			this.label7.Text = "Banco";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEntidadFinanciera
			// 
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance102;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 208;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEntidadFinanciera.Enabled = false;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(64, 40);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(160, 21);
			this.cmbEntidadFinanciera.TabIndex = 752;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
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
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance103;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.pgEntrega);
			this.utcReportes.Controls.Add(this.utpEstado);
			this.utcReportes.Controls.Add(this.ultraTabPageControl1);
			this.utcReportes.Controls.Add(this.ultraTabPageControl2);
			this.utcReportes.Location = new System.Drawing.Point(8, 72);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(1112, 424);
			this.utcReportes.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcReportes.TabIndex = 753;
			ultraTab1.TabPage = this.utpEstado;
			ultraTab1.Text = "Estado - Saldos - Ratios - Variaciones";
			ultraTab2.TabPage = this.pgEntrega;
			ultraTab2.Text = "Evolución Cartera Vencida - Primeras Cuotas";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Resumen de Colocaciones - Análisis de Cosechas";
			ultraTab4.TabPage = this.ultraTabPageControl2;
			ultraTab4.Text = "Cuotas Pagadas de Creditos de Retiro y Fondo de Garantia";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3,
																																													ultraTab4});
			this.utcReportes.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1110, 403);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(240, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 754;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// frmCre_ReportesEstadisticosCartera
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1128, 502);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.utcReportes);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmCre_ReportesEstadisticosCartera";
			this.Text = "Estadisticos De Cartera";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ReportesEstadisticosCartera_KeyDown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmCre_ReportesEstadisticosCartera_KeyPress);
			this.Load += new System.EventHandler(this.frmCre_ReportesEstadisticosCartera_Load);
			this.utpEstado.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCoberturaFDG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFDG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatioVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatio30)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo30)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCartera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridPrimerasCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGtEvolucionRiesgo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridAnalisisdeCosechas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumenColocaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridFDG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetiros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCre_ReportesEstadisticosCartera_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082902");
			if (!Funcion.Permiso("861", cdsSeteoF))
			{				
				MessageBox.Show("No tiene acceso a Estadistica de Cartera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("24/04/2018");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;

			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera Where idEntidadFinanciera In (1, 3, 4) Order By idEntidadFinanciera");
			this.cmbEntidadFinanciera.Value = 4;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
//			int idBanco = 0;
//			if (this.cmbEntidadFinanciera.ActiveRow != null) idBanco = (int)this.cmbEntidadFinanciera.Value;
//
//			ParameterFields paramFields = new ParameterFields ();
//
//			ParameterField paramFieldFecha = new ParameterField ();
//			ParameterDiscreteValue discreteValFecha= new ParameterDiscreteValue ();
//			paramFieldFecha.ParameterFieldName = "@Fecha";
//			discreteValFecha.Value = (DateTime)this.dtFecha.Value;
//			paramFieldFecha.CurrentValues.Add (discreteValFecha);
//			paramFields.Add (paramFieldFecha);
//
//			ParameterField pfCapital = new ParameterField ();
//			ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
//			pfCapital.ParameterFieldName = "@idEntidadFinanciera";
//			ValCapital.Value = idBanco;
//			pfCapital.CurrentValues.Add (ValCapital);
//			paramFields.Add (pfCapital);
//			
//			Reporte miRep = new Reporte("ReportesEstadisticosCartera.rpt", "");
//			miRep.MdiParent = this.MdiParent;			
//			miRep.crVista.ParameterFieldInfo = paramFields;
//			miRep.Show();
		}

		private void DiseñoGridTurno(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bIniciaMes, bool bSumatorias, string sColumna)
		{
			try
			{				
//				this.uGridCartera.DisplayLayout.Bands[0].Columns["Numero"].Hidden = true;			
				this.uGridCartera.DisplayLayout.Bands[0].Columns["RangoMorosidad"].Width = 120;
				if (this.uGridCartera.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridCartera.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";

					for (int j = 2; j < iContColumnas; j++)
					{
						nombrCol = this.uGridCartera.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
						this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						if(nombrCol =="Operaciones")
						{
							this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0";
							this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
						}
						else
						{
							this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
							this.uGridCartera.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
						}
					}
				}

				#region Diseño Grid
				if (uGrid.Rows.Count > 0)
				{
					int iColIni = 2;

				
					if (uGrid.Rows.Count > 0)
					{
						string nombrCol = "";
							
						int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
						for (int j = iColIni; j < iColumnas; j++)
						{
							
							nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;			
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
	
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
							uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
							
						}					
					}				
				}	
				#endregion Diseño Grid
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridCartera.DataSource = null;	
			string sSQL = string.Format("Exec ReporteEstadosPoint '{0}', {1}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridCartera);

			#region Ratio
			this.txtSaldoVencido.Value = 0.00m;
			this.txtRatioVencido.Value = 0.00m;
			this.txtSaldo30.Value = 0.00m;
			this.txtRatio30.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCartera.Rows.All)
			{				
				if((int)dr.Cells["Numero"].Value > 3)
				{
					this.txtSaldoVencido.Value = Convert.ToDecimal(this.txtSaldoVencido.Value) + Convert.ToDecimal(dr.Cells["VencidoTotal"].Value);
					this.txtRatioVencido.Value = Convert.ToDecimal(this.txtRatioVencido.Value) + Convert.ToDecimal(dr.Cells["ParticipacionSaldo"].Value);					
				}

				if((int)dr.Cells["Numero"].Value > 4)
				{
					this.txtSaldo30.Value = Convert.ToDecimal(this.txtSaldo30.Value) + Convert.ToDecimal(dr.Cells["VencidoTotal"].Value);
					this.txtRatio30.Value = Convert.ToDecimal(this.txtRatio30.Value) + Convert.ToDecimal(dr.Cells["ParticipacionSaldo"].Value);
				}				
			}
			#endregion Ratio

			#region Saldo FDG
			this.txtSaldoFDG.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec Cre_SaldoFondoGarantia '{0}', {1}, 0", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value));
			this.txtCoberturaFDG.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec Cre_SaldoFondoGarantia '{0}', {1}, 1", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value));
			#endregion Saldo FDG

			#region Variaciones
			DateTime dtInicio = DateTime.Parse("01/" + Convert.ToDateTime(this.dtFecha.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtFecha.Value).Year.ToString());

			string sSQLRpt = string.Format("Exec Cre_ReporteVariacionesPoint '{0}','{1}'", dtInicio.ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLRpt, this.uGridVariaciones);
			#endregion Variaciones

			#region consultas

			string sSQLEVR = string.Format("Exec Cre_ReporteRiesgoTotal '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLEVR, this.uGtEvolucionRiesgo);

			this.uGridPrimerasCuotas.DataSource = null;

			this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Header.Caption = "Primeras Cuotas";

			string sSQLPC = string.Format("Exec Cre_TablaPrimerasCuotas '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLPC, this.uGridPrimerasCuotas);

			DiseñoGrid();

			string sSQLRC = string.Format("Exec Cre_ResumenColocaciones '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLRC, this.uGridResumenColocaciones);

			string sSQLAC = string.Format("Exec Cre_AnalisisCosechas '{0}', {1}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value);
			FuncionesProcedimientos.dsGeneral(sSQLAC, this.uGridAnalisisdeCosechas);

			this.uGridRetiros.DataSource = null;

			this.uGridRetiros.DisplayLayout.Bands[0].Header.Caption = "Cuotas De Retiros De Mercaderia";

			string sSQLCR = string.Format("Exec Cre_ReporteCuotasPagadasRetiro '{0}'", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLCR, this.uGridRetiros);

			DiseñoGridCuotas(this.uGridRetiros);

			this.uGridFDG.DataSource = null;

			this.uGridFDG.DisplayLayout.Bands[0].Header.Caption = "Cuotas De Fondo De Garantia";

			string sSQLCFDG = string.Format("Exec Cre_ReporteCuotasPagadasFDG '{0}'", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQLCFDG, this.uGridFDG);

			DiseñoGridCuotas(this.uGridFDG);
				#endregion consultas

					this.DiseñoGridTurno(this.uGridCartera, true, false, "");


		}

		private void DiseñoGridCuotas(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			if (uGrid.Rows.Count > 0)
			{
				uGrid.DisplayLayout.Bands[0].Columns["Periodo"].Width = 60;
				
				uGrid.DisplayLayout.Bands[0].Columns["Periodo"].CellActivation = Activation.Disabled;
								
				uGrid.DisplayLayout.Bands[0].Columns["Periodo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
								
				int iContColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
				string nombrCol = "";

				for (int j = 1; j < iContColumnas; j++)
				{
					nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 40;
					uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0";

					uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
				}
			}
		}

		private void DiseñoGrid()
		{
			if (this.uGridPrimerasCuotas.Rows.Count > 0)
			{
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Fecha"].Width = 80;
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Lote"].Width = 60;

				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.Disabled;
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Lote"].CellActivation = Activation.Disabled;
				
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Fecha"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Lote"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
								
				this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns["Lote"].Format = "#,##0";
				
				int iContColumnas = (int)this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns.Count;
				
				string nombrCol = "";

				for (int j = 2; j < iContColumnas; j++)
				{
					nombrCol = this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[nombrCol].Width = 50;
					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0";

					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridPrimerasCuotas.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
				}
			}
		}

		private void utpEstado_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void uGridCartera_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (e.Cell.Column.ToString() == sColumna)
			{
//				int Numcuotas = Convert.ToInt32(sColumna);
								
				using (frmCre_ReportesEstadisticosCarteraDetalle CCJ = new frmCre_ReportesEstadisticosCarteraDetalle ((int) this.cmbEntidadFinanciera.Value,(DateTime) this.dtFecha.Value, (int)e.Cell.Row.Cells["Numero"].Value))
				{
					if (DialogResult.OK == CCJ.ShowDialog())
					{

					}
				}			 				
			}
		}

		private void frmCre_ReportesEstadisticosCartera_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

		}

		private void frmCre_ReportesEstadisticosCartera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();	
		}
	}
}
