using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoAnual.
	/// </summary>
	public class frmPresupuestoAnual : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasPrecios;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasMargen;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOtrosIngresos;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadTotal;
		private System.Windows.Forms.Button btnProyeccion;
		private System.Windows.Forms.Label label17;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoAnual()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoAnual));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 16, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 16, true);
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 4, true);
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 5, true);
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 6, true);
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 7, true);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 8, true);
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 9, true);
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 10, true);
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 11, true);
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 12, true);
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 14, true);
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 16, true);
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 17, true);
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 5, true);
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 6, true);
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 7, true);
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 8, true);
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 9, true);
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 10, true);
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 11, true);
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 12, true);
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 14, true);
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 16, true);
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 17, true);
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVentasPrecios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridVentasMargen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridOtrosIngresos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtGastosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtUtilidadTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.btnProyeccion = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasMargen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtrosIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.Decimal);
			ultraDataColumn2.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn3.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn14});
			// 
			// uGridVentas
			// 
			this.uGridVentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridVentas.DisplayLayout.Appearance = appearance1;
			this.uGridVentas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 139;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 93;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 82;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 75;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 76;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 76;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 76;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 76;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance10;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 77;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 78;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance12;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 79;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 100;
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
																										 ultraGridColumn14});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentas.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.CellAppearance = appearance17;
			this.uGridVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentas.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentas.Location = new System.Drawing.Point(152, 0);
			this.uGridVentas.Name = "uGridVentas";
			this.uGridVentas.Size = new System.Drawing.Size(1081, 40);
			this.uGridVentas.TabIndex = 17;
			this.uGridVentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentas_KeyDown);
			this.uGridVentas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVentas_AfterCellUpdate);
			this.uGridVentas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentas_InitializeLayout);
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(8, 8);
			this.txtAño.MinValue = 2017;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 739;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.Location = new System.Drawing.Point(112, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(30, 23);
			this.btnVer.TabIndex = 740;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// ultraDataSource2
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
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
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn27.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn28.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn29.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn30.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// uGridVentasPrecios
			// 
			this.uGridVentasPrecios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasPrecios.DataSource = this.ultraDataSource2;
			appearance22.BackColor = System.Drawing.Color.White;
			this.uGridVentasPrecios.DisplayLayout.Appearance = appearance22;
			this.uGridVentasPrecios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.Caption = "Año";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 137;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "Venta";
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 98;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance23;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "%";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 44;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance24;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 78;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance25;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 78;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance26;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Width = 78;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance27;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 78;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance28;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 78;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance29;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Width = 78;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance30;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Width = 78;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance31;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 11;
			ultraGridColumn26.Width = 80;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance32;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Width = 80;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance33;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 13;
			ultraGridColumn28.Width = 80;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance34;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 14;
			ultraGridColumn29.Width = 80;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance35;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 15;
			ultraGridColumn30.Width = 80;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance36;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance37;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance38;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance39;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance40;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance41;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance42;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance43;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance44;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance45;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance46;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance47;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance48;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance49;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13});
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridVentasPrecios.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.ActiveRowAppearance = appearance50;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance51.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.CardAreaAppearance = appearance51;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.CellAppearance = appearance52;
			this.uGridVentasPrecios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance53.FontData.BoldAsString = "True";
			appearance53.FontData.Name = "Arial";
			appearance53.FontData.SizeInPoints = 8F;
			appearance53.ForeColor = System.Drawing.Color.White;
			appearance53.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.HeaderAppearance = appearance53;
			appearance54.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance54.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowAlternateAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowSelectorAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.SelectedRowAppearance = appearance56;
			this.uGridVentasPrecios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasPrecios.Location = new System.Drawing.Point(8, 40);
			this.uGridVentasPrecios.Name = "uGridVentasPrecios";
			this.uGridVentasPrecios.Size = new System.Drawing.Size(1226, 112);
			this.uGridVentasPrecios.TabIndex = 742;
			this.uGridVentasPrecios.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridVentasPrecios_KeyPress);
			this.uGridVentasPrecios.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridVentasPrecios_BeforeCellUpdate);
			this.uGridVentasPrecios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasPrecios_KeyDown);
			this.uGridVentasPrecios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVentasPrecios_AfterCellUpdate);
			this.uGridVentasPrecios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasPrecios_InitializeLayout);
			// 
			// uGridVentasMargen
			// 
			this.uGridVentasMargen.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasMargen.DataSource = this.ultraDataSource3;
			appearance57.BackColor = System.Drawing.Color.White;
			this.uGridVentasMargen.DisplayLayout.Appearance = appearance57;
			this.uGridVentasMargen.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.Caption = "Año";
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 137;
			ultraGridColumn33.Header.VisiblePosition = 16;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Header.Caption = "Margen";
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 93;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance58;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "%";
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Width = 40;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance59;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 80;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance60;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance61;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 5;
			ultraGridColumn38.Width = 80;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance62;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance63;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 7;
			ultraGridColumn40.Width = 80;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance64;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 8;
			ultraGridColumn41.Width = 80;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance65;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 9;
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance66;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 10;
			ultraGridColumn43.Width = 80;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance67;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 11;
			ultraGridColumn44.Width = 80;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance68;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 12;
			ultraGridColumn45.Width = 80;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance69;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 13;
			ultraGridColumn46.Width = 80;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance70;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 14;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance71;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 15;
			ultraGridColumn48.Width = 95;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
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
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance72;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance73;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance74;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance75;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance76;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance77;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance78;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance79;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance80;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance81;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance82;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance83;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance84;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26});
			ultraGridBand3.SummaryFooterCaption = "TOTALES";
			this.uGridVentasMargen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance85.ForeColor = System.Drawing.Color.Black;
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasMargen.DisplayLayout.Override.ActiveRowAppearance = appearance85;
			this.uGridVentasMargen.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasMargen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasMargen.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance86.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasMargen.DisplayLayout.Override.CardAreaAppearance = appearance86;
			appearance87.ForeColor = System.Drawing.Color.Black;
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasMargen.DisplayLayout.Override.CellAppearance = appearance87;
			this.uGridVentasMargen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance88.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance88.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance88.FontData.BoldAsString = "True";
			appearance88.FontData.Name = "Arial";
			appearance88.FontData.SizeInPoints = 8F;
			appearance88.ForeColor = System.Drawing.Color.White;
			appearance88.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasMargen.DisplayLayout.Override.HeaderAppearance = appearance88;
			appearance89.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance89.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance89.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.RowAlternateAppearance = appearance89;
			appearance90.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance90.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.RowSelectorAppearance = appearance90;
			appearance91.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance91.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance91.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.SelectedRowAppearance = appearance91;
			this.uGridVentasMargen.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasMargen.Location = new System.Drawing.Point(8, 152);
			this.uGridVentasMargen.Name = "uGridVentasMargen";
			this.uGridVentasMargen.Size = new System.Drawing.Size(1226, 112);
			this.uGridVentasMargen.TabIndex = 743;
			this.uGridVentasMargen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasMargen_KeyDown);
			this.uGridVentasMargen.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVentasMargen_AfterCellUpdate);
			this.uGridVentasMargen.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasMargen_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn36.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn39.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn40.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn41.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn42.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn48.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
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
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48});
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.Location = new System.Drawing.Point(8, 591);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(30, 23);
			this.btnNuevo.TabIndex = 744;
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn53.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn54.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn55.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn57.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn58.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn59.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn60.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn61.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn62.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn63.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn64.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn65.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn66.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66});
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource4;
			appearance92.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance92;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn49.Header.Caption = "Año";
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 137;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn51.Header.Caption = "Gasto";
			ultraGridColumn51.Header.VisiblePosition = 2;
			ultraGridColumn51.Width = 97;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance93;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "%";
			ultraGridColumn52.Header.VisiblePosition = 3;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 40;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance94;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 4;
			ultraGridColumn53.Width = 80;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance95;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 5;
			ultraGridColumn54.Width = 81;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance96;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 6;
			ultraGridColumn55.Width = 81;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance97;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 7;
			ultraGridColumn56.Width = 81;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance98;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 8;
			ultraGridColumn57.Width = 81;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance99;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 9;
			ultraGridColumn58.Width = 81;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance100;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 10;
			ultraGridColumn59.Width = 81;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance101;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 11;
			ultraGridColumn60.Width = 81;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance102;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 12;
			ultraGridColumn61.Width = 81;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance103;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 13;
			ultraGridColumn62.Width = 81;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance104;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 14;
			ultraGridColumn63.Width = 81;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance105;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 15;
			ultraGridColumn64.Width = 81;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance106;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 16;
			ultraGridColumn65.Width = 81;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance107;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 17;
			ultraGridColumn66.Width = 96;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66});
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance108;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance109;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance110;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance111;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance112;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance113;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance114;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance115;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance116;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance117;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance118;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance119;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance120.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance120;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance121.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance121;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings27,
																																															summarySettings28,
																																															summarySettings29,
																																															summarySettings30,
																																															summarySettings31,
																																															summarySettings32,
																																															summarySettings33,
																																															summarySettings34,
																																															summarySettings35,
																																															summarySettings36,
																																															summarySettings37,
																																															summarySettings38,
																																															summarySettings39,
																																															summarySettings40});
			ultraGridBand4.SummaryFooterCaption = "TOTALES";
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance122.ForeColor = System.Drawing.Color.Black;
			appearance122.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance122;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance123.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance123;
			appearance124.ForeColor = System.Drawing.Color.Black;
			appearance124.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance124;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance125.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance125.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance125.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance125.FontData.BoldAsString = "True";
			appearance125.FontData.Name = "Arial";
			appearance125.FontData.SizeInPoints = 8F;
			appearance125.ForeColor = System.Drawing.Color.White;
			appearance125.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance125;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance126.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance126.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance126.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance126;
			appearance127.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance127.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance127.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance127;
			appearance128.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance128.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance128.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance128;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 376);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1226, 168);
			this.uGridGastos.TabIndex = 745;
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn67.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn72.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn73.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn74.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn75.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn76.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn77.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn78.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn79.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn80.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn81.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn82.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn83.DataType = typeof(System.Decimal);
			ultraDataColumn83.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn84.DataType = typeof(System.Decimal);
			ultraDataColumn84.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81,
																																 ultraDataColumn82,
																																 ultraDataColumn83,
																																 ultraDataColumn84});
			// 
			// uGridOtrosIngresos
			// 
			this.uGridOtrosIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOtrosIngresos.DataSource = this.ultraDataSource5;
			appearance129.BackColor = System.Drawing.Color.White;
			this.uGridOtrosIngresos.DisplayLayout.Appearance = appearance129;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn67.Header.Caption = "Año";
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 137;
			ultraGridColumn68.Header.VisiblePosition = 17;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn69.Header.Caption = "Otros Ingresos";
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridColumn69.Width = 93;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance130;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.Caption = "%";
			ultraGridColumn70.Header.VisiblePosition = 2;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 40;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance131;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 3;
			ultraGridColumn71.Width = 80;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance132;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 4;
			ultraGridColumn72.Width = 80;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance133;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 5;
			ultraGridColumn73.Width = 80;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance134;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 6;
			ultraGridColumn74.Width = 80;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance135;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 7;
			ultraGridColumn75.Width = 80;
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance136;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 8;
			ultraGridColumn76.Width = 80;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance137;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 9;
			ultraGridColumn77.Width = 80;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance138;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 10;
			ultraGridColumn78.Width = 80;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance139;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 11;
			ultraGridColumn79.Width = 80;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance140;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 12;
			ultraGridColumn80.Width = 80;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance141;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 13;
			ultraGridColumn81.Width = 80;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance142;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 14;
			ultraGridColumn82.Width = 80;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance143;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.VisiblePosition = 15;
			ultraGridColumn83.Width = 80;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance144;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 16;
			ultraGridColumn84.Width = 95;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84});
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance145;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance146;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance147;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance148;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance149;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance150;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance151;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance152;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance153;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance154;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance155.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance155;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance156.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance156;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance157.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance157;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings41,
																																															summarySettings42,
																																															summarySettings43,
																																															summarySettings44,
																																															summarySettings45,
																																															summarySettings46,
																																															summarySettings47,
																																															summarySettings48,
																																															summarySettings49,
																																															summarySettings50,
																																															summarySettings51,
																																															summarySettings52,
																																															summarySettings53});
			ultraGridBand5.SummaryFooterCaption = "TOTALES";
			this.uGridOtrosIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance158.ForeColor = System.Drawing.Color.Black;
			appearance158.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOtrosIngresos.DisplayLayout.Override.ActiveRowAppearance = appearance158;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance159.BackColor = System.Drawing.Color.Transparent;
			this.uGridOtrosIngresos.DisplayLayout.Override.CardAreaAppearance = appearance159;
			appearance160.ForeColor = System.Drawing.Color.Black;
			appearance160.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOtrosIngresos.DisplayLayout.Override.CellAppearance = appearance160;
			this.uGridOtrosIngresos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance161.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance161.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance161.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance161.FontData.BoldAsString = "True";
			appearance161.FontData.Name = "Arial";
			appearance161.FontData.SizeInPoints = 8F;
			appearance161.ForeColor = System.Drawing.Color.White;
			appearance161.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOtrosIngresos.DisplayLayout.Override.HeaderAppearance = appearance161;
			appearance162.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance162.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance162.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.RowAlternateAppearance = appearance162;
			appearance163.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance163.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance163.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.RowSelectorAppearance = appearance163;
			appearance164.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance164.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance164.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.SelectedRowAppearance = appearance164;
			this.uGridOtrosIngresos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOtrosIngresos.Location = new System.Drawing.Point(8, 264);
			this.uGridOtrosIngresos.Name = "uGridOtrosIngresos";
			this.uGridOtrosIngresos.Size = new System.Drawing.Size(1226, 112);
			this.uGridOtrosIngresos.TabIndex = 746;
			this.uGridOtrosIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridOtrosIngresos_KeyDown);
			this.uGridOtrosIngresos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOtrosIngresos_AfterCellUpdate);
			this.uGridOtrosIngresos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOtrosIngresos_InitializeLayout);
			// 
			// txtGastosEnero
			// 
			this.txtGastosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance165.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosEnero.Appearance = appearance165;
			this.txtGastosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosEnero.Enabled = false;
			this.txtGastosEnero.FormatString = "#,##0.00";
			this.txtGastosEnero.Location = new System.Drawing.Point(80, 592);
			this.txtGastosEnero.Name = "txtGastosEnero";
			this.txtGastosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosEnero.PromptChar = ' ';
			this.txtGastosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosEnero.TabIndex = 750;
			// 
			// txtGastosFebrero
			// 
			this.txtGastosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance166.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosFebrero.Appearance = appearance166;
			this.txtGastosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosFebrero.Enabled = false;
			this.txtGastosFebrero.FormatString = "#,##0.00";
			this.txtGastosFebrero.Location = new System.Drawing.Point(168, 592);
			this.txtGastosFebrero.Name = "txtGastosFebrero";
			this.txtGastosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFebrero.PromptChar = ' ';
			this.txtGastosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosFebrero.TabIndex = 749;
			// 
			// txtIngresosFebrero
			// 
			this.txtIngresosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance167.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosFebrero.Appearance = appearance167;
			this.txtIngresosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosFebrero.Enabled = false;
			this.txtIngresosFebrero.FormatString = "#,##0.00";
			this.txtIngresosFebrero.Location = new System.Drawing.Point(168, 568);
			this.txtIngresosFebrero.Name = "txtIngresosFebrero";
			this.txtIngresosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosFebrero.PromptChar = ' ';
			this.txtIngresosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosFebrero.TabIndex = 748;
			// 
			// txtIngresosEnero
			// 
			this.txtIngresosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance168.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosEnero.Appearance = appearance168;
			this.txtIngresosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosEnero.Enabled = false;
			this.txtIngresosEnero.FormatString = "#,##0.00";
			this.txtIngresosEnero.Location = new System.Drawing.Point(80, 568);
			this.txtIngresosEnero.Name = "txtIngresosEnero";
			this.txtIngresosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosEnero.PromptChar = ' ';
			this.txtIngresosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosEnero.TabIndex = 747;
			// 
			// txtGastosMarzo
			// 
			this.txtGastosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance169.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMarzo.Appearance = appearance169;
			this.txtGastosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMarzo.Enabled = false;
			this.txtGastosMarzo.FormatString = "#,##0.00";
			this.txtGastosMarzo.Location = new System.Drawing.Point(256, 592);
			this.txtGastosMarzo.Name = "txtGastosMarzo";
			this.txtGastosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMarzo.PromptChar = ' ';
			this.txtGastosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMarzo.TabIndex = 754;
			// 
			// txtGastosAbril
			// 
			this.txtGastosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance170.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAbril.Appearance = appearance170;
			this.txtGastosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAbril.Enabled = false;
			this.txtGastosAbril.FormatString = "#,##0.00";
			this.txtGastosAbril.Location = new System.Drawing.Point(344, 592);
			this.txtGastosAbril.Name = "txtGastosAbril";
			this.txtGastosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAbril.PromptChar = ' ';
			this.txtGastosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAbril.TabIndex = 753;
			// 
			// txtIngresosAbril
			// 
			this.txtIngresosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance171.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAbril.Appearance = appearance171;
			this.txtIngresosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAbril.Enabled = false;
			this.txtIngresosAbril.FormatString = "#,##0.00";
			this.txtIngresosAbril.Location = new System.Drawing.Point(344, 568);
			this.txtIngresosAbril.Name = "txtIngresosAbril";
			this.txtIngresosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAbril.PromptChar = ' ';
			this.txtIngresosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAbril.TabIndex = 752;
			// 
			// txtIngresosMarzo
			// 
			this.txtIngresosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance172.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMarzo.Appearance = appearance172;
			this.txtIngresosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMarzo.Enabled = false;
			this.txtIngresosMarzo.FormatString = "#,##0.00";
			this.txtIngresosMarzo.Location = new System.Drawing.Point(256, 568);
			this.txtIngresosMarzo.Name = "txtIngresosMarzo";
			this.txtIngresosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMarzo.PromptChar = ' ';
			this.txtIngresosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMarzo.TabIndex = 751;
			// 
			// txtGastosJulio
			// 
			this.txtGastosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance173.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJulio.Appearance = appearance173;
			this.txtGastosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJulio.Enabled = false;
			this.txtGastosJulio.FormatString = "#,##0.00";
			this.txtGastosJulio.Location = new System.Drawing.Point(608, 592);
			this.txtGastosJulio.Name = "txtGastosJulio";
			this.txtGastosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJulio.PromptChar = ' ';
			this.txtGastosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJulio.TabIndex = 762;
			// 
			// txtGastosAgosto
			// 
			this.txtGastosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance174.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAgosto.Appearance = appearance174;
			this.txtGastosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAgosto.Enabled = false;
			this.txtGastosAgosto.FormatString = "#,##0.00";
			this.txtGastosAgosto.Location = new System.Drawing.Point(696, 592);
			this.txtGastosAgosto.Name = "txtGastosAgosto";
			this.txtGastosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAgosto.PromptChar = ' ';
			this.txtGastosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAgosto.TabIndex = 761;
			// 
			// txtIngresosAgosto
			// 
			this.txtIngresosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance175.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAgosto.Appearance = appearance175;
			this.txtIngresosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAgosto.Enabled = false;
			this.txtIngresosAgosto.FormatString = "#,##0.00";
			this.txtIngresosAgosto.Location = new System.Drawing.Point(696, 568);
			this.txtIngresosAgosto.Name = "txtIngresosAgosto";
			this.txtIngresosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAgosto.PromptChar = ' ';
			this.txtIngresosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAgosto.TabIndex = 760;
			// 
			// txtIngresosJulio
			// 
			this.txtIngresosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance176.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJulio.Appearance = appearance176;
			this.txtIngresosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJulio.Enabled = false;
			this.txtIngresosJulio.FormatString = "#,##0.00";
			this.txtIngresosJulio.Location = new System.Drawing.Point(608, 568);
			this.txtIngresosJulio.Name = "txtIngresosJulio";
			this.txtIngresosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJulio.PromptChar = ' ';
			this.txtIngresosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJulio.TabIndex = 759;
			// 
			// txtGastosMayo
			// 
			this.txtGastosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance177.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMayo.Appearance = appearance177;
			this.txtGastosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMayo.Enabled = false;
			this.txtGastosMayo.FormatString = "#,##0.00";
			this.txtGastosMayo.Location = new System.Drawing.Point(432, 592);
			this.txtGastosMayo.Name = "txtGastosMayo";
			this.txtGastosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMayo.PromptChar = ' ';
			this.txtGastosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMayo.TabIndex = 758;
			// 
			// txtGastosJunio
			// 
			this.txtGastosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance178.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJunio.Appearance = appearance178;
			this.txtGastosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJunio.Enabled = false;
			this.txtGastosJunio.FormatString = "#,##0.00";
			this.txtGastosJunio.Location = new System.Drawing.Point(520, 592);
			this.txtGastosJunio.Name = "txtGastosJunio";
			this.txtGastosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJunio.PromptChar = ' ';
			this.txtGastosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJunio.TabIndex = 757;
			// 
			// txtIngresosJunio
			// 
			this.txtIngresosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance179.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJunio.Appearance = appearance179;
			this.txtIngresosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJunio.Enabled = false;
			this.txtIngresosJunio.FormatString = "#,##0.00";
			this.txtIngresosJunio.Location = new System.Drawing.Point(520, 568);
			this.txtIngresosJunio.Name = "txtIngresosJunio";
			this.txtIngresosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJunio.PromptChar = ' ';
			this.txtIngresosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJunio.TabIndex = 756;
			// 
			// txtIngresosMayo
			// 
			this.txtIngresosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance180.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMayo.Appearance = appearance180;
			this.txtIngresosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMayo.Enabled = false;
			this.txtIngresosMayo.FormatString = "#,##0.00";
			this.txtIngresosMayo.Location = new System.Drawing.Point(432, 568);
			this.txtIngresosMayo.Name = "txtIngresosMayo";
			this.txtIngresosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMayo.PromptChar = ' ';
			this.txtIngresosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMayo.TabIndex = 755;
			// 
			// txtGastosNoviembre
			// 
			this.txtGastosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance181.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosNoviembre.Appearance = appearance181;
			this.txtGastosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosNoviembre.Enabled = false;
			this.txtGastosNoviembre.FormatString = "#,##0.00";
			this.txtGastosNoviembre.Location = new System.Drawing.Point(960, 592);
			this.txtGastosNoviembre.Name = "txtGastosNoviembre";
			this.txtGastosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosNoviembre.PromptChar = ' ';
			this.txtGastosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosNoviembre.TabIndex = 770;
			// 
			// txtGastosDiciembre
			// 
			this.txtGastosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance182.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosDiciembre.Appearance = appearance182;
			this.txtGastosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosDiciembre.Enabled = false;
			this.txtGastosDiciembre.FormatString = "#,##0.00";
			this.txtGastosDiciembre.Location = new System.Drawing.Point(1048, 592);
			this.txtGastosDiciembre.Name = "txtGastosDiciembre";
			this.txtGastosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosDiciembre.PromptChar = ' ';
			this.txtGastosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosDiciembre.TabIndex = 769;
			// 
			// txtIngresosDiciembre
			// 
			this.txtIngresosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance183.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosDiciembre.Appearance = appearance183;
			this.txtIngresosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosDiciembre.Enabled = false;
			this.txtIngresosDiciembre.FormatString = "#,##0.00";
			this.txtIngresosDiciembre.Location = new System.Drawing.Point(1048, 568);
			this.txtIngresosDiciembre.Name = "txtIngresosDiciembre";
			this.txtIngresosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosDiciembre.PromptChar = ' ';
			this.txtIngresosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosDiciembre.TabIndex = 768;
			// 
			// txtIngresosNoviembre
			// 
			this.txtIngresosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance184.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosNoviembre.Appearance = appearance184;
			this.txtIngresosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosNoviembre.Enabled = false;
			this.txtIngresosNoviembre.FormatString = "#,##0.00";
			this.txtIngresosNoviembre.Location = new System.Drawing.Point(960, 568);
			this.txtIngresosNoviembre.Name = "txtIngresosNoviembre";
			this.txtIngresosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosNoviembre.PromptChar = ' ';
			this.txtIngresosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosNoviembre.TabIndex = 767;
			// 
			// txtGastosSeptiembre
			// 
			this.txtGastosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance185.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosSeptiembre.Appearance = appearance185;
			this.txtGastosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosSeptiembre.Enabled = false;
			this.txtGastosSeptiembre.FormatString = "#,##0.00";
			this.txtGastosSeptiembre.Location = new System.Drawing.Point(784, 592);
			this.txtGastosSeptiembre.Name = "txtGastosSeptiembre";
			this.txtGastosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosSeptiembre.PromptChar = ' ';
			this.txtGastosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosSeptiembre.TabIndex = 766;
			// 
			// txtGastosOctubre
			// 
			this.txtGastosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance186.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosOctubre.Appearance = appearance186;
			this.txtGastosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosOctubre.Enabled = false;
			this.txtGastosOctubre.FormatString = "#,##0.00";
			this.txtGastosOctubre.Location = new System.Drawing.Point(872, 592);
			this.txtGastosOctubre.Name = "txtGastosOctubre";
			this.txtGastosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosOctubre.PromptChar = ' ';
			this.txtGastosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosOctubre.TabIndex = 765;
			// 
			// txtIngresosOctubre
			// 
			this.txtIngresosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance187.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosOctubre.Appearance = appearance187;
			this.txtIngresosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosOctubre.Enabled = false;
			this.txtIngresosOctubre.FormatString = "#,##0.00";
			this.txtIngresosOctubre.Location = new System.Drawing.Point(872, 568);
			this.txtIngresosOctubre.Name = "txtIngresosOctubre";
			this.txtIngresosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosOctubre.PromptChar = ' ';
			this.txtIngresosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosOctubre.TabIndex = 764;
			// 
			// txtIngresosSeptiembre
			// 
			this.txtIngresosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance188.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosSeptiembre.Appearance = appearance188;
			this.txtIngresosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosSeptiembre.Enabled = false;
			this.txtIngresosSeptiembre.FormatString = "#,##0.00";
			this.txtIngresosSeptiembre.Location = new System.Drawing.Point(784, 568);
			this.txtIngresosSeptiembre.Name = "txtIngresosSeptiembre";
			this.txtIngresosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosSeptiembre.PromptChar = ' ';
			this.txtIngresosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosSeptiembre.TabIndex = 763;
			// 
			// txtUtilidadNoviembre
			// 
			this.txtUtilidadNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance189.ForeColor = System.Drawing.Color.Firebrick;
			appearance189.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadNoviembre.Appearance = appearance189;
			this.txtUtilidadNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadNoviembre.Enabled = false;
			this.txtUtilidadNoviembre.FormatString = "#,##0.00";
			this.txtUtilidadNoviembre.Location = new System.Drawing.Point(960, 624);
			this.txtUtilidadNoviembre.Name = "txtUtilidadNoviembre";
			this.txtUtilidadNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadNoviembre.PromptChar = ' ';
			this.txtUtilidadNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadNoviembre.TabIndex = 782;
			// 
			// txtUtilidadDiciembre
			// 
			this.txtUtilidadDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance190.ForeColor = System.Drawing.Color.Firebrick;
			appearance190.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadDiciembre.Appearance = appearance190;
			this.txtUtilidadDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadDiciembre.Enabled = false;
			this.txtUtilidadDiciembre.FormatString = "#,##0.00";
			this.txtUtilidadDiciembre.Location = new System.Drawing.Point(1048, 624);
			this.txtUtilidadDiciembre.Name = "txtUtilidadDiciembre";
			this.txtUtilidadDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadDiciembre.PromptChar = ' ';
			this.txtUtilidadDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadDiciembre.TabIndex = 781;
			// 
			// txtUtilidadSeptiembre
			// 
			this.txtUtilidadSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance191.ForeColor = System.Drawing.Color.Firebrick;
			appearance191.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadSeptiembre.Appearance = appearance191;
			this.txtUtilidadSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadSeptiembre.Enabled = false;
			this.txtUtilidadSeptiembre.FormatString = "#,##0.00";
			this.txtUtilidadSeptiembre.Location = new System.Drawing.Point(784, 624);
			this.txtUtilidadSeptiembre.Name = "txtUtilidadSeptiembre";
			this.txtUtilidadSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadSeptiembre.PromptChar = ' ';
			this.txtUtilidadSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadSeptiembre.TabIndex = 780;
			// 
			// txtUtilidadOctubre
			// 
			this.txtUtilidadOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance192.ForeColor = System.Drawing.Color.Firebrick;
			appearance192.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadOctubre.Appearance = appearance192;
			this.txtUtilidadOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadOctubre.Enabled = false;
			this.txtUtilidadOctubre.FormatString = "#,##0.00";
			this.txtUtilidadOctubre.Location = new System.Drawing.Point(872, 624);
			this.txtUtilidadOctubre.Name = "txtUtilidadOctubre";
			this.txtUtilidadOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadOctubre.PromptChar = ' ';
			this.txtUtilidadOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadOctubre.TabIndex = 779;
			// 
			// txtUtilidadJulio
			// 
			this.txtUtilidadJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance193.ForeColor = System.Drawing.Color.Firebrick;
			appearance193.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJulio.Appearance = appearance193;
			this.txtUtilidadJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJulio.Enabled = false;
			this.txtUtilidadJulio.FormatString = "#,##0.00";
			this.txtUtilidadJulio.Location = new System.Drawing.Point(608, 624);
			this.txtUtilidadJulio.Name = "txtUtilidadJulio";
			this.txtUtilidadJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJulio.PromptChar = ' ';
			this.txtUtilidadJulio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJulio.TabIndex = 778;
			// 
			// txtUtilidadAgosto
			// 
			this.txtUtilidadAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance194.ForeColor = System.Drawing.Color.Firebrick;
			appearance194.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAgosto.Appearance = appearance194;
			this.txtUtilidadAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAgosto.Enabled = false;
			this.txtUtilidadAgosto.FormatString = "#,##0.00";
			this.txtUtilidadAgosto.Location = new System.Drawing.Point(696, 624);
			this.txtUtilidadAgosto.Name = "txtUtilidadAgosto";
			this.txtUtilidadAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAgosto.PromptChar = ' ';
			this.txtUtilidadAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAgosto.TabIndex = 777;
			// 
			// txtUtilidadMayo
			// 
			this.txtUtilidadMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance195.ForeColor = System.Drawing.Color.Firebrick;
			appearance195.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMayo.Appearance = appearance195;
			this.txtUtilidadMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMayo.Enabled = false;
			this.txtUtilidadMayo.FormatString = "#,##0.00";
			this.txtUtilidadMayo.Location = new System.Drawing.Point(432, 624);
			this.txtUtilidadMayo.Name = "txtUtilidadMayo";
			this.txtUtilidadMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMayo.PromptChar = ' ';
			this.txtUtilidadMayo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMayo.TabIndex = 776;
			// 
			// txtUtilidadJunio
			// 
			this.txtUtilidadJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance196.ForeColor = System.Drawing.Color.Firebrick;
			appearance196.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJunio.Appearance = appearance196;
			this.txtUtilidadJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJunio.Enabled = false;
			this.txtUtilidadJunio.FormatString = "#,##0.00";
			this.txtUtilidadJunio.Location = new System.Drawing.Point(520, 624);
			this.txtUtilidadJunio.Name = "txtUtilidadJunio";
			this.txtUtilidadJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJunio.PromptChar = ' ';
			this.txtUtilidadJunio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJunio.TabIndex = 775;
			// 
			// txtUtilidadMarzo
			// 
			this.txtUtilidadMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance197.ForeColor = System.Drawing.Color.Firebrick;
			appearance197.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMarzo.Appearance = appearance197;
			this.txtUtilidadMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMarzo.Enabled = false;
			this.txtUtilidadMarzo.FormatString = "#,##0.00";
			this.txtUtilidadMarzo.Location = new System.Drawing.Point(256, 624);
			this.txtUtilidadMarzo.Name = "txtUtilidadMarzo";
			this.txtUtilidadMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMarzo.PromptChar = ' ';
			this.txtUtilidadMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMarzo.TabIndex = 774;
			// 
			// txtUtilidadAbril
			// 
			this.txtUtilidadAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance198.ForeColor = System.Drawing.Color.Firebrick;
			appearance198.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAbril.Appearance = appearance198;
			this.txtUtilidadAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAbril.Enabled = false;
			this.txtUtilidadAbril.FormatString = "#,##0.00";
			this.txtUtilidadAbril.Location = new System.Drawing.Point(344, 624);
			this.txtUtilidadAbril.Name = "txtUtilidadAbril";
			this.txtUtilidadAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAbril.PromptChar = ' ';
			this.txtUtilidadAbril.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAbril.TabIndex = 773;
			// 
			// txtUtilidadEnero
			// 
			this.txtUtilidadEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance199.ForeColor = System.Drawing.Color.Firebrick;
			appearance199.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadEnero.Appearance = appearance199;
			this.txtUtilidadEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadEnero.Enabled = false;
			this.txtUtilidadEnero.FormatString = "#,##0.00";
			this.txtUtilidadEnero.Location = new System.Drawing.Point(80, 624);
			this.txtUtilidadEnero.Name = "txtUtilidadEnero";
			this.txtUtilidadEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadEnero.PromptChar = ' ';
			this.txtUtilidadEnero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadEnero.TabIndex = 772;
			// 
			// txtUtilidadFebrero
			// 
			this.txtUtilidadFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance200.ForeColor = System.Drawing.Color.Firebrick;
			appearance200.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadFebrero.Appearance = appearance200;
			this.txtUtilidadFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadFebrero.Enabled = false;
			this.txtUtilidadFebrero.FormatString = "#,##0.00";
			this.txtUtilidadFebrero.Location = new System.Drawing.Point(168, 624);
			this.txtUtilidadFebrero.Name = "txtUtilidadFebrero";
			this.txtUtilidadFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadFebrero.PromptChar = ' ';
			this.txtUtilidadFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadFebrero.TabIndex = 771;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(40, 570);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 783;
			this.label5.Text = "(+) I";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(40, 594);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 784;
			this.label1.Text = "(-) G";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(40, 626);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 785;
			this.label2.Text = "(=) U";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUtilidadTotal
			// 
			this.txtUtilidadTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance201.FontData.BoldAsString = "True";
			appearance201.ForeColor = System.Drawing.Color.Firebrick;
			appearance201.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotal.Appearance = appearance201;
			this.txtUtilidadTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotal.Enabled = false;
			this.txtUtilidadTotal.FormatString = "#,##0.00";
			this.txtUtilidadTotal.Location = new System.Drawing.Point(1136, 624);
			this.txtUtilidadTotal.Name = "txtUtilidadTotal";
			this.txtUtilidadTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotal.PromptChar = ' ';
			this.txtUtilidadTotal.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotal.TabIndex = 788;
			// 
			// txtGastosTotal
			// 
			this.txtGastosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance202.FontData.BoldAsString = "True";
			appearance202.ForeColor = System.Drawing.Color.Firebrick;
			appearance202.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGastosTotal.Appearance = appearance202;
			this.txtGastosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosTotal.Enabled = false;
			this.txtGastosTotal.FormatString = "#,##0.00";
			this.txtGastosTotal.Location = new System.Drawing.Point(1136, 592);
			this.txtGastosTotal.Name = "txtGastosTotal";
			this.txtGastosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosTotal.PromptChar = ' ';
			this.txtGastosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtGastosTotal.TabIndex = 787;
			// 
			// txtIngresosTotal
			// 
			this.txtIngresosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance203.FontData.BoldAsString = "True";
			appearance203.ForeColor = System.Drawing.Color.Firebrick;
			appearance203.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtIngresosTotal.Appearance = appearance203;
			this.txtIngresosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosTotal.Enabled = false;
			this.txtIngresosTotal.FormatString = "#,##0.00";
			this.txtIngresosTotal.Location = new System.Drawing.Point(1136, 568);
			this.txtIngresosTotal.Name = "txtIngresosTotal";
			this.txtIngresosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosTotal.PromptChar = ' ';
			this.txtIngresosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtIngresosTotal.TabIndex = 786;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(80, 546);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 789;
			this.label3.Text = "ENERO";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(168, 546);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 790;
			this.label4.Text = "FEBRERO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(256, 546);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 791;
			this.label6.Text = "MARZO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(344, 546);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 792;
			this.label7.Text = "ABRIL";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(696, 546);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 796;
			this.label8.Text = "AGOSTO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(608, 546);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 795;
			this.label9.Text = "JULIO";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(520, 546);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 794;
			this.label10.Text = "JUNIO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(432, 546);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 793;
			this.label11.Text = "MAYO";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(1048, 546);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(80, 16);
			this.label12.TabIndex = 800;
			this.label12.Text = "DICIEMBRE";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(960, 546);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 799;
			this.label13.Text = "NOVIEMBRE";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(872, 546);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(80, 16);
			this.label14.TabIndex = 798;
			this.label14.Text = "OCTUBRE";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(784, 546);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 797;
			this.label15.Text = "SEPTIEMBRE";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Firebrick;
			this.label16.Location = new System.Drawing.Point(1136, 546);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(100, 16);
			this.label16.TabIndex = 801;
			this.label16.Text = "TOTAL";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnProyeccion
			// 
			this.btnProyeccion.CausesValidation = false;
			this.btnProyeccion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnProyeccion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnProyeccion.Image = ((System.Drawing.Image)(resources.GetObject("btnProyeccion.Image")));
			this.btnProyeccion.Location = new System.Drawing.Point(8, 567);
			this.btnProyeccion.Name = "btnProyeccion";
			this.btnProyeccion.Size = new System.Drawing.Size(30, 23);
			this.btnProyeccion.TabIndex = 803;
			this.btnProyeccion.Click += new System.EventHandler(this.btnProyeccion_Click);
			// 
			// label17
			// 
			this.label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label17.Location = new System.Drawing.Point(0, 618);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(1240, 1);
			this.label17.TabIndex = 804;
			// 
			// frmPresupuestoAnual
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1240, 654);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.btnProyeccion);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtUtilidadTotal);
			this.Controls.Add(this.txtGastosTotal);
			this.Controls.Add(this.txtIngresosTotal);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtUtilidadNoviembre);
			this.Controls.Add(this.txtUtilidadDiciembre);
			this.Controls.Add(this.txtUtilidadSeptiembre);
			this.Controls.Add(this.txtUtilidadOctubre);
			this.Controls.Add(this.txtUtilidadJulio);
			this.Controls.Add(this.txtUtilidadAgosto);
			this.Controls.Add(this.txtUtilidadMayo);
			this.Controls.Add(this.txtUtilidadJunio);
			this.Controls.Add(this.txtUtilidadMarzo);
			this.Controls.Add(this.txtUtilidadAbril);
			this.Controls.Add(this.txtUtilidadEnero);
			this.Controls.Add(this.txtUtilidadFebrero);
			this.Controls.Add(this.txtGastosNoviembre);
			this.Controls.Add(this.txtGastosDiciembre);
			this.Controls.Add(this.txtIngresosDiciembre);
			this.Controls.Add(this.txtIngresosNoviembre);
			this.Controls.Add(this.txtGastosSeptiembre);
			this.Controls.Add(this.txtGastosOctubre);
			this.Controls.Add(this.txtIngresosOctubre);
			this.Controls.Add(this.txtIngresosSeptiembre);
			this.Controls.Add(this.txtGastosJulio);
			this.Controls.Add(this.txtGastosAgosto);
			this.Controls.Add(this.txtIngresosAgosto);
			this.Controls.Add(this.txtIngresosJulio);
			this.Controls.Add(this.txtGastosMayo);
			this.Controls.Add(this.txtGastosJunio);
			this.Controls.Add(this.txtIngresosJunio);
			this.Controls.Add(this.txtIngresosMayo);
			this.Controls.Add(this.txtGastosMarzo);
			this.Controls.Add(this.txtGastosAbril);
			this.Controls.Add(this.txtIngresosAbril);
			this.Controls.Add(this.txtIngresosMarzo);
			this.Controls.Add(this.txtGastosEnero);
			this.Controls.Add(this.txtGastosFebrero);
			this.Controls.Add(this.txtIngresosFebrero);
			this.Controls.Add(this.txtIngresosEnero);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridVentas);
			this.Controls.Add(this.uGridVentasPrecios);
			this.Controls.Add(this.uGridVentasMargen);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.uGridOtrosIngresos);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoAnual";
			this.Text = "Presupuesto Anual";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoAnual_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoAnual_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasMargen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtrosIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaPresupuestoVentas {0}", (int)this.txtAño.Value);
			this.uGridVentas.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.SumatoriasAlConsultar(this.uGridVentas);

			sSQL = string.Format("Exec ConsultaPresupuestoVentasPrecios {0}", (int)this.txtAño.Value);
			this.uGridVentasPrecios.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);	
			
			this.SumatoriasAlConsultar(this.uGridVentasPrecios);

			sSQL = string.Format("Exec ConsultaPresupuestoVentasMargen {0}", (int)this.txtAño.Value);
			this.uGridVentasMargen.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);	

			this.SumatoriasAlConsultar(this.uGridVentasMargen);

			sSQL = string.Format("Exec ConsultaPresupuestoOtrosIngresos {0}", (int)this.txtAño.Value);
			this.uGridOtrosIngresos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);				

			this.SumatoriasAlConsultar(this.uGridOtrosIngresos);

			sSQL = string.Format("Exec ConsultaPresupuestoGastos {0}", (int)this.txtAño.Value);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);				

			this.SumatoriasAlConsultar(this.uGridGastos);

			this.Resultado();
		}

		private void frmPresupuestoAnual_Load(object sender, System.EventArgs e)
		{
			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
		}

		private void uGridVentasPrecios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void uGridVentas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void uGridVentasMargen_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void uGridVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentas);
		}

		private void uGridVentasPrecios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentasPrecios);
		}

		private void uGridVentasMargen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentasMargen);
		}

		private void ActualizaValores (int iMes, int idTipoPrecio, decimal dValor, int iTabla)
		{
			string sSQL = string.Format("Exec PresupuestoActualizaInformacion {0}, {1}, {2}, {3}, {4}", 
				(int)this.txtAño.Value, iMes, idTipoPrecio, dValor, iTabla);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ActualizaPorcentajes (int idTipoPrecio, decimal dPorcentaje, int iTabla)
		{
			string sSQL = string.Format("Exec PresupuestoActualizaPorcentaje {0}, {1}, {2}, {3}", 
				(int)this.txtAño.Value, idTipoPrecio, dPorcentaje, iTabla);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void Resultado ()
		{
			this.txtIngresosEnero.Value = 0.00m;
			this.txtIngresosFebrero.Value = 0.00m;
			this.txtIngresosMarzo.Value = 0.00m;
			this.txtIngresosAbril.Value = 0.00m;
			this.txtIngresosMayo.Value = 0.00m;
			this.txtIngresosJunio.Value = 0.00m;
			this.txtIngresosJulio.Value = 0.00m;
			this.txtIngresosAgosto.Value = 0.00m;
			this.txtIngresosSeptiembre.Value = 0.00m;
			this.txtIngresosOctubre.Value = 0.00m;
			this.txtIngresosNoviembre.Value = 0.00m;
			this.txtIngresosDiciembre.Value = 0.00m;
			this.txtIngresosTotal.Value = 0.00m;
	
			this.txtGastosEnero.Value = 0.00m;
			this.txtGastosFebrero.Value = 0.00m;
			this.txtGastosMarzo.Value = 0.00m;
			this.txtGastosAbril.Value = 0.00m;
			this.txtGastosMayo.Value = 0.00m;
			this.txtGastosJunio.Value = 0.00m;
			this.txtGastosJulio.Value = 0.00m;
			this.txtGastosAgosto.Value = 0.00m;
			this.txtGastosSeptiembre.Value = 0.00m;
			this.txtGastosOctubre.Value = 0.00m;
			this.txtGastosNoviembre.Value = 0.00m;
			this.txtGastosDiciembre.Value = 0.00m;
			this.txtGastosTotal.Value = 0.00m;
	
			this.txtUtilidadEnero.Value = 0.00m;
			this.txtUtilidadFebrero.Value = 0.00m;
			this.txtUtilidadMarzo.Value = 0.00m;
			this.txtUtilidadAbril.Value = 0.00m;
			this.txtUtilidadMayo.Value = 0.00m;
			this.txtUtilidadJunio.Value = 0.00m;
			this.txtUtilidadJulio.Value = 0.00m;
			this.txtUtilidadAgosto.Value = 0.00m;
			this.txtUtilidadSeptiembre.Value = 0.00m;
			this.txtUtilidadOctubre.Value = 0.00m;
			this.txtUtilidadNoviembre.Value = 0.00m;
			this.txtUtilidadDiciembre.Value = 0.00m;
			this.txtUtilidadTotal.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasMargen.Rows.All)
			{
				this.txtIngresosEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) + Convert.ToDecimal(dr.Cells["ENERO"].Value);
				this.txtIngresosFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) + Convert.ToDecimal(dr.Cells["FEBRERO"].Value);
				this.txtIngresosMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) + Convert.ToDecimal(dr.Cells["MARZO"].Value);
				this.txtIngresosAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) + Convert.ToDecimal(dr.Cells["ABRIL"].Value);
				this.txtIngresosMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) + Convert.ToDecimal(dr.Cells["MAYO"].Value);
				this.txtIngresosJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) + Convert.ToDecimal(dr.Cells["JUNIO"].Value);
				this.txtIngresosJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				this.txtIngresosAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				this.txtIngresosSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				this.txtIngresosOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				this.txtIngresosNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				this.txtIngresosDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtIngresosTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOtrosIngresos.Rows.All)
			{
				this.txtIngresosEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) + Convert.ToDecimal(dr.Cells["ENERO"].Value);
				this.txtIngresosFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) + Convert.ToDecimal(dr.Cells["FEBRERO"].Value);
				this.txtIngresosMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) + Convert.ToDecimal(dr.Cells["MARZO"].Value);
				this.txtIngresosAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) + Convert.ToDecimal(dr.Cells["ABRIL"].Value);
				this.txtIngresosMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) + Convert.ToDecimal(dr.Cells["MAYO"].Value);
				this.txtIngresosJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) + Convert.ToDecimal(dr.Cells["JUNIO"].Value);
				this.txtIngresosJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				this.txtIngresosAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				this.txtIngresosSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				this.txtIngresosOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				this.txtIngresosNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				this.txtIngresosDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtIngresosTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				this.txtGastosEnero.Value = Convert.ToDecimal(this.txtGastosEnero.Value) + Convert.ToDecimal(dr.Cells["ENERO"].Value);
				this.txtGastosFebrero.Value = Convert.ToDecimal(this.txtGastosFebrero.Value) + Convert.ToDecimal(dr.Cells["FEBRERO"].Value);
				this.txtGastosMarzo.Value = Convert.ToDecimal(this.txtGastosMarzo.Value) + Convert.ToDecimal(dr.Cells["MARZO"].Value);
				this.txtGastosAbril.Value = Convert.ToDecimal(this.txtGastosAbril.Value) + Convert.ToDecimal(dr.Cells["ABRIL"].Value);
				this.txtGastosMayo.Value = Convert.ToDecimal(this.txtGastosMayo.Value) + Convert.ToDecimal(dr.Cells["MAYO"].Value);
				this.txtGastosJunio.Value = Convert.ToDecimal(this.txtGastosJunio.Value) + Convert.ToDecimal(dr.Cells["JUNIO"].Value);
				this.txtGastosJulio.Value = Convert.ToDecimal(this.txtGastosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				this.txtGastosAgosto.Value = Convert.ToDecimal(this.txtGastosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				this.txtGastosSeptiembre.Value = Convert.ToDecimal(this.txtGastosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				this.txtGastosOctubre.Value = Convert.ToDecimal(this.txtGastosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				this.txtGastosNoviembre.Value = Convert.ToDecimal(this.txtGastosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				this.txtGastosDiciembre.Value = Convert.ToDecimal(this.txtGastosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtGastosTotal.Value = Convert.ToDecimal(this.txtGastosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			this.txtUtilidadEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) - Convert.ToDecimal(this.txtGastosEnero.Value);
			this.txtUtilidadFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) - Convert.ToDecimal(this.txtGastosFebrero.Value);
			this.txtUtilidadMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) - Convert.ToDecimal(this.txtGastosMarzo.Value);
			this.txtUtilidadAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) - Convert.ToDecimal(this.txtGastosAbril.Value);
			this.txtUtilidadMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) - Convert.ToDecimal(this.txtGastosMayo.Value);
			this.txtUtilidadJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) - Convert.ToDecimal(this.txtGastosJunio.Value);
			this.txtUtilidadJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) - Convert.ToDecimal(this.txtGastosJulio.Value);
			this.txtUtilidadAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) - Convert.ToDecimal(this.txtGastosAgosto.Value);
			this.txtUtilidadSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) - Convert.ToDecimal(this.txtGastosSeptiembre.Value);
			this.txtUtilidadOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) - Convert.ToDecimal(this.txtGastosOctubre.Value);
			this.txtUtilidadNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) - Convert.ToDecimal(this.txtGastosNoviembre.Value);
			this.txtUtilidadDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) - Convert.ToDecimal(this.txtGastosDiciembre.Value);
			this.txtUtilidadTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) - Convert.ToDecimal(this.txtGastosTotal.Value);	
		}

		private void SumatoriasAlConsultar(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["ENERO"].Value) + 
					Convert.ToDecimal(dr.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(dr.Cells["MARZO"].Value) +
					Convert.ToDecimal(dr.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(dr.Cells["MAYO"].Value) +
					Convert.ToDecimal(dr.Cells["JUNIO"].Value) +
					Convert.ToDecimal(dr.Cells["JULIO"].Value) + 
					Convert.ToDecimal(dr.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
			}
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["TOTAL"].Value = 
				Convert.ToDecimal(e.Cell.Row.Cells["ENERO"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["FEBRERO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["MARZO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["ABRIL"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["MAYO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["JUNIO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["JULIO"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["AGOSTO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["SEPTIEMBRE"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["OCTUBRE"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["NOVIEMBRE"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["DICIEMBRE"].Value);			
		}
				
		private void uGridVentas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (1, 0, Convert.ToDecimal(e.Cell.Value), 1);
				
				ActualizaValoresVentas("ENERO", e);

				Sumatorias(e);				
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (2, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("FEBRERO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (3, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("MARZO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (4, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("ABRIL", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (5, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("MAYO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (6, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("JUNIO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (7, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("JULIO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (8, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("AGOSTO", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (9, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("SEPTIEMBRE", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (10, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("OCTUBRE", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (11, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("NOVIEMBRE", e);

				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (12, 0, Convert.ToDecimal(e.Cell.Value), 1);

				ActualizaValoresVentas("DICIEMBRE", e);

				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void ActualizaValoresVentas(string sColumna, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
			{
				dr.Cells[sColumna].Value = Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value) * Convert.ToDecimal(dr.Cells["Porcentaje"].Value) / 100;	
			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.uGridVentasMargen.Rows.All)
				{
					if ((int)dr.Cells["idTipoPrecio"].Value == (int)dr1.Cells["idTipoPrecio"].Value)
						dr1.Cells[sColumna].Value = Convert.ToDecimal(dr.Cells[sColumna].Value) * Convert.ToDecimal(dr1.Cells["Porcentaje"].Value) / 100;						
				}
			}
		}


		private void ActualizaValoresMargen(string sColumna, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{										
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasMargen.Rows.All)
			{
				if ((int)dr.Cells["idTipoPrecio"].Value == (int)e.Cell.Row.Cells["idTipoPrecio"].Value)
					dr.Cells[sColumna].Value = Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value) * Convert.ToDecimal(dr.Cells["Porcentaje"].Value) / 100;				
			}		
		}

		private void ActualizaValoresComisiones(string sColumna)
		{
			decimal dVentasAlmacenes = 0.00m;
			decimal dVentasCredito = 0.00m;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
			{
				if ((int)dr.Cells["idTipoPrecio"].Value > 1)
					dVentasAlmacenes = dVentasAlmacenes + Convert.ToDecimal(dr.Cells[sColumna].Value);

				if ((int)dr.Cells["idTipoPrecio"].Value == 4)
					dVentasCredito = Convert.ToDecimal(dr.Cells[sColumna].Value);
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.uGridGastos.Rows.All)
			{
				if ((int)dr1.Cells["idCuentaAgrupa"].Value == 16) //LOCALES
					dr1.Cells[sColumna].Value = Decimal.Round((dVentasAlmacenes * 1.20m / 100.00m) + 14761.15m, 2);
				
				if ((int)dr1.Cells["idCuentaAgrupa"].Value == 21) //CREDITO
					dr1.Cells[sColumna].Value = Decimal.Round((dVentasCredito * 0.60m / 100.00m), 2);
			}
		}

		private void uGridVentasPrecios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				e.Cell.Row.Cells["ENERO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["ENERO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["FEBRERO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["FEBRERO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["MARZO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["MARZO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["ABRIL"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["ABRIL"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["MAYO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["MAYO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["JUNIO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["JUNIO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["JULIO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["JULIO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["AGOSTO"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["AGOSTO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["SEPTIEMBRE"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["SEPTIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["OCTUBRE"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["OCTUBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["NOVIEMBRE"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["NOVIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
				e.Cell.Row.Cells["DICIEMBRE"].Value = Convert.ToDecimal(this.uGridVentas.ActiveRow.Cells["DICIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;

				ActualizaPorcentajes((int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				
			}

			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (1, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("ENERO", e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (2, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("FEBRERO", e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (3, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("MARZO", e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (4, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("ABRIL", e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (5, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("MAYO", e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (6, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("JUNIO", e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (7, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("JULIO", e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (8, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("AGOSTO", e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (9, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("SEPTIEMBRE", e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (10, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("OCTUBRE", e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (11, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("NOVIEMBRE", e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (12, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 2);				

				ActualizaValoresMargen("DICIEMBRE", e);
			}
			#endregion Actualiza Valores

			Sumatorias(e);
		}

		private void uGridVentasMargen_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
				{
					if ((int)dr.Cells["idTipoPrecio"].Value == (int)e.Cell.Row.Cells["idTipoPrecio"].Value)
					{
						e.Cell.Row.Cells["ENERO"].Value = Convert.ToDecimal(dr.Cells["ENERO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["FEBRERO"].Value = Convert.ToDecimal(dr.Cells["FEBRERO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["MARZO"].Value = Convert.ToDecimal(dr.Cells["MARZO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["ABRIL"].Value = Convert.ToDecimal(dr.Cells["ABRIL"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["MAYO"].Value = Convert.ToDecimal(dr.Cells["MAYO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["JUNIO"].Value = Convert.ToDecimal(dr.Cells["JUNIO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["JULIO"].Value = Convert.ToDecimal(dr.Cells["JULIO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["AGOSTO"].Value = Convert.ToDecimal(dr.Cells["AGOSTO"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["SEPTIEMBRE"].Value = Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["OCTUBRE"].Value = Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["NOVIEMBRE"].Value = Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
						e.Cell.Row.Cells["DICIEMBRE"].Value = Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value) / 100;
					}
				}

				ActualizaPorcentajes((int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);
			}

			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (1, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);	
			
				ActualizaValoresComisiones("ENERO");
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (2, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("FEBRERO");
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (3, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("MARZO");
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (4, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("ABRIL");
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (5, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("MAYO");
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (6, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("JUNIO");
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (7, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("JULIO");
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (8, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("AGOSTO");
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (9, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("SEPTIEMBRE");
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (10, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("OCTUBRE");
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (11, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("NOVIEMBRE");
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (12, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, Convert.ToDecimal(e.Cell.Value), 3);				

				ActualizaValoresComisiones("DICIEMBRE");
			}
			#endregion Actualiza Valores

			Sumatorias(e);
		}

		private void uGridVentasPrecios_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Porcentaje")
				if ((int)e.Cell.Row.Cells["idTipoPrecio"].Value == 1) e.Cancel = true;
		}

		private void frmPresupuestoAnual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec PresupuestoGeneraNuevoPeriodo {0}", (int)this.txtAño.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void uGridVentasPrecios_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.uGridVentasPrecios.ActiveRow.Cells["idTipoPrecio"].Value > 1 && 
				this.uGridVentasPrecios.ActiveCell.Column.Key.ToString().ToUpper() != "PORCENTAJE") e.Handled = true;
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGastos);
		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Valor")
			{				
				e.Cell.Row.Cells["ENERO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["FEBRERO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["MARZO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["ABRIL"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["MAYO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["JUNIO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["JULIO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["AGOSTO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["SEPTIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["OCTUBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["NOVIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["DICIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
			
				ActualizaPorcentajes((int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);
			}

			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (1, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (2, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (3, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (4, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (5, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (6, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (7, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (8, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (9, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (10, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (11, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (12, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, Convert.ToDecimal(e.Cell.Value), 4);				
			}
			#endregion Actualiza Valores

			Sumatorias(e);
		}

		private void uGridOtrosIngresos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void uGridOtrosIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridOtrosIngresos);
		}

		private void uGridOtrosIngresos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Valor")
			{				
				e.Cell.Row.Cells["ENERO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["FEBRERO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["MARZO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["ABRIL"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["MAYO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["JUNIO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["JULIO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["AGOSTO"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["SEPTIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["OCTUBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["NOVIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
				e.Cell.Row.Cells["DICIEMBRE"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);
			
				ActualizaPorcentajes((int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);
			}

			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (1, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (2, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (3, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (4, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (5, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (6, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (7, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (8, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (9, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (10, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (11, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (12, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDecimal(e.Cell.Value), 5);				
			}
			#endregion Actualiza Valores

			Sumatorias(e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnProyeccion_Click(object sender, System.EventArgs e)
		{
			frmProyeccionPresupuesto PP = new frmProyeccionPresupuesto((int)this.txtAño.Value);
			PP.Show();
//			frmPresupuestoProyeccion PP1 = new frmPresupuestoProyeccion((int)this.txtAño.Value);
//			PP1.Show();
		}		
	}
}
