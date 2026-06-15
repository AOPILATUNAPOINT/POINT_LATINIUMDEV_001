using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoRealDentalis.
	/// </summary>
	public class frmPresupuestoRealDentalis : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosTotal;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label5;
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
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosEnero;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasPrecios;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoRealDentalis()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int iYear = 0;
		public frmPresupuestoRealDentalis(int IYear)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			 iYear = IYear;
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
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 2, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 3, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 4, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 5, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 6, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 7, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 8, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 9, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 10, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 11, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 14, true);
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 2, true);
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 3, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 4, true);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 5, true);
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 6, true);
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 7, true);
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 8, true);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 9, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 10, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 11, true);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 14, true);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoRealDentalis));
			this.label17 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtUtilidadTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
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
			this.txtGastosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVentasPrecios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).BeginInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label17
			// 
			this.label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label17.Location = new System.Drawing.Point(16, 429);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(1240, 1);
			this.label17.TabIndex = 924;
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Firebrick;
			this.label16.Location = new System.Drawing.Point(1120, 357);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(100, 16);
			this.label16.TabIndex = 922;
			this.label16.Text = "TOTAL";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(1032, 357);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(80, 16);
			this.label12.TabIndex = 921;
			this.label12.Text = "DICIEMBRE";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(944, 357);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 920;
			this.label13.Text = "NOVIEMBRE";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(856, 357);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(80, 16);
			this.label14.TabIndex = 919;
			this.label14.Text = "OCTUBRE";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(768, 357);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 918;
			this.label15.Text = "SEPTIEMBRE";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(680, 357);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 917;
			this.label8.Text = "AGOSTO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(592, 357);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 916;
			this.label9.Text = "JULIO";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(504, 357);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 915;
			this.label10.Text = "JUNIO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(416, 357);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 914;
			this.label11.Text = "MAYO";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(328, 357);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 913;
			this.label7.Text = "ABRIL";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(240, 357);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 912;
			this.label6.Text = "MARZO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(152, 357);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 911;
			this.label4.Text = "FEBRERO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(64, 357);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 910;
			this.label3.Text = "ENERO";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtUtilidadTotal
			// 
			this.txtUtilidadTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.FontData.BoldAsString = "True";
			appearance1.ForeColor = System.Drawing.Color.Firebrick;
			appearance1.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotal.Appearance = appearance1;
			this.txtUtilidadTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotal.Enabled = false;
			this.txtUtilidadTotal.FormatString = "#,##0.00";
			this.txtUtilidadTotal.Location = new System.Drawing.Point(1120, 437);
			this.txtUtilidadTotal.Name = "txtUtilidadTotal";
			this.txtUtilidadTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotal.PromptChar = ' ';
			this.txtUtilidadTotal.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotal.TabIndex = 909;
			// 
			// txtGastosTotal
			// 
			this.txtGastosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.FontData.BoldAsString = "True";
			appearance2.ForeColor = System.Drawing.Color.Firebrick;
			appearance2.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGastosTotal.Appearance = appearance2;
			this.txtGastosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosTotal.Enabled = false;
			this.txtGastosTotal.FormatString = "#,##0.00";
			this.txtGastosTotal.Location = new System.Drawing.Point(1120, 405);
			this.txtGastosTotal.Name = "txtGastosTotal";
			this.txtGastosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosTotal.PromptChar = ' ';
			this.txtGastosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtGastosTotal.TabIndex = 908;
			// 
			// txtIngresosTotal
			// 
			this.txtIngresosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance3.FontData.BoldAsString = "True";
			appearance3.ForeColor = System.Drawing.Color.Firebrick;
			appearance3.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtIngresosTotal.Appearance = appearance3;
			this.txtIngresosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosTotal.Enabled = false;
			this.txtIngresosTotal.FormatString = "#,##0.00";
			this.txtIngresosTotal.Location = new System.Drawing.Point(1120, 381);
			this.txtIngresosTotal.Name = "txtIngresosTotal";
			this.txtIngresosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosTotal.PromptChar = ' ';
			this.txtIngresosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtIngresosTotal.TabIndex = 907;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(24, 437);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 906;
			this.label2.Text = "(=) U";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(24, 405);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 905;
			this.label1.Text = "(-) G";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(24, 381);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 904;
			this.label5.Text = "(+) I";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUtilidadNoviembre
			// 
			this.txtUtilidadNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance4.ForeColor = System.Drawing.Color.Firebrick;
			appearance4.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadNoviembre.Appearance = appearance4;
			this.txtUtilidadNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadNoviembre.Enabled = false;
			this.txtUtilidadNoviembre.FormatString = "#,##0.00";
			this.txtUtilidadNoviembre.Location = new System.Drawing.Point(944, 437);
			this.txtUtilidadNoviembre.Name = "txtUtilidadNoviembre";
			this.txtUtilidadNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadNoviembre.PromptChar = ' ';
			this.txtUtilidadNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadNoviembre.TabIndex = 903;
			// 
			// txtUtilidadDiciembre
			// 
			this.txtUtilidadDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			appearance5.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadDiciembre.Appearance = appearance5;
			this.txtUtilidadDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadDiciembre.Enabled = false;
			this.txtUtilidadDiciembre.FormatString = "#,##0.00";
			this.txtUtilidadDiciembre.Location = new System.Drawing.Point(1032, 437);
			this.txtUtilidadDiciembre.Name = "txtUtilidadDiciembre";
			this.txtUtilidadDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadDiciembre.PromptChar = ' ';
			this.txtUtilidadDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadDiciembre.TabIndex = 902;
			// 
			// txtUtilidadSeptiembre
			// 
			this.txtUtilidadSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.ForeColor = System.Drawing.Color.Firebrick;
			appearance6.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadSeptiembre.Appearance = appearance6;
			this.txtUtilidadSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadSeptiembre.Enabled = false;
			this.txtUtilidadSeptiembre.FormatString = "#,##0.00";
			this.txtUtilidadSeptiembre.Location = new System.Drawing.Point(768, 437);
			this.txtUtilidadSeptiembre.Name = "txtUtilidadSeptiembre";
			this.txtUtilidadSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadSeptiembre.PromptChar = ' ';
			this.txtUtilidadSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadSeptiembre.TabIndex = 901;
			// 
			// txtUtilidadOctubre
			// 
			this.txtUtilidadOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.ForeColor = System.Drawing.Color.Firebrick;
			appearance7.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadOctubre.Appearance = appearance7;
			this.txtUtilidadOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadOctubre.Enabled = false;
			this.txtUtilidadOctubre.FormatString = "#,##0.00";
			this.txtUtilidadOctubre.Location = new System.Drawing.Point(856, 437);
			this.txtUtilidadOctubre.Name = "txtUtilidadOctubre";
			this.txtUtilidadOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadOctubre.PromptChar = ' ';
			this.txtUtilidadOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadOctubre.TabIndex = 900;
			// 
			// txtUtilidadJulio
			// 
			this.txtUtilidadJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColor = System.Drawing.Color.Firebrick;
			appearance8.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJulio.Appearance = appearance8;
			this.txtUtilidadJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJulio.Enabled = false;
			this.txtUtilidadJulio.FormatString = "#,##0.00";
			this.txtUtilidadJulio.Location = new System.Drawing.Point(592, 437);
			this.txtUtilidadJulio.Name = "txtUtilidadJulio";
			this.txtUtilidadJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJulio.PromptChar = ' ';
			this.txtUtilidadJulio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJulio.TabIndex = 899;
			// 
			// txtUtilidadAgosto
			// 
			this.txtUtilidadAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColor = System.Drawing.Color.Firebrick;
			appearance9.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAgosto.Appearance = appearance9;
			this.txtUtilidadAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAgosto.Enabled = false;
			this.txtUtilidadAgosto.FormatString = "#,##0.00";
			this.txtUtilidadAgosto.Location = new System.Drawing.Point(680, 437);
			this.txtUtilidadAgosto.Name = "txtUtilidadAgosto";
			this.txtUtilidadAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAgosto.PromptChar = ' ';
			this.txtUtilidadAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAgosto.TabIndex = 898;
			// 
			// txtUtilidadMayo
			// 
			this.txtUtilidadMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColor = System.Drawing.Color.Firebrick;
			appearance10.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMayo.Appearance = appearance10;
			this.txtUtilidadMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMayo.Enabled = false;
			this.txtUtilidadMayo.FormatString = "#,##0.00";
			this.txtUtilidadMayo.Location = new System.Drawing.Point(416, 437);
			this.txtUtilidadMayo.Name = "txtUtilidadMayo";
			this.txtUtilidadMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMayo.PromptChar = ' ';
			this.txtUtilidadMayo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMayo.TabIndex = 897;
			// 
			// txtUtilidadJunio
			// 
			this.txtUtilidadJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance11.ForeColor = System.Drawing.Color.Firebrick;
			appearance11.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJunio.Appearance = appearance11;
			this.txtUtilidadJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJunio.Enabled = false;
			this.txtUtilidadJunio.FormatString = "#,##0.00";
			this.txtUtilidadJunio.Location = new System.Drawing.Point(504, 437);
			this.txtUtilidadJunio.Name = "txtUtilidadJunio";
			this.txtUtilidadJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJunio.PromptChar = ' ';
			this.txtUtilidadJunio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJunio.TabIndex = 896;
			// 
			// txtUtilidadMarzo
			// 
			this.txtUtilidadMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance12.ForeColor = System.Drawing.Color.Firebrick;
			appearance12.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMarzo.Appearance = appearance12;
			this.txtUtilidadMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMarzo.Enabled = false;
			this.txtUtilidadMarzo.FormatString = "#,##0.00";
			this.txtUtilidadMarzo.Location = new System.Drawing.Point(240, 437);
			this.txtUtilidadMarzo.Name = "txtUtilidadMarzo";
			this.txtUtilidadMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMarzo.PromptChar = ' ';
			this.txtUtilidadMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMarzo.TabIndex = 895;
			// 
			// txtUtilidadAbril
			// 
			this.txtUtilidadAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.ForeColor = System.Drawing.Color.Firebrick;
			appearance13.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAbril.Appearance = appearance13;
			this.txtUtilidadAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAbril.Enabled = false;
			this.txtUtilidadAbril.FormatString = "#,##0.00";
			this.txtUtilidadAbril.Location = new System.Drawing.Point(328, 437);
			this.txtUtilidadAbril.Name = "txtUtilidadAbril";
			this.txtUtilidadAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAbril.PromptChar = ' ';
			this.txtUtilidadAbril.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAbril.TabIndex = 894;
			// 
			// txtUtilidadEnero
			// 
			this.txtUtilidadEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance14.ForeColor = System.Drawing.Color.Firebrick;
			appearance14.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadEnero.Appearance = appearance14;
			this.txtUtilidadEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadEnero.Enabled = false;
			this.txtUtilidadEnero.FormatString = "#,##0.00";
			this.txtUtilidadEnero.Location = new System.Drawing.Point(64, 437);
			this.txtUtilidadEnero.Name = "txtUtilidadEnero";
			this.txtUtilidadEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadEnero.PromptChar = ' ';
			this.txtUtilidadEnero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadEnero.TabIndex = 893;
			// 
			// txtUtilidadFebrero
			// 
			this.txtUtilidadFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColor = System.Drawing.Color.Firebrick;
			appearance15.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadFebrero.Appearance = appearance15;
			this.txtUtilidadFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadFebrero.Enabled = false;
			this.txtUtilidadFebrero.FormatString = "#,##0.00";
			this.txtUtilidadFebrero.Location = new System.Drawing.Point(152, 437);
			this.txtUtilidadFebrero.Name = "txtUtilidadFebrero";
			this.txtUtilidadFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadFebrero.PromptChar = ' ';
			this.txtUtilidadFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadFebrero.TabIndex = 892;
			// 
			// txtGastosNoviembre
			// 
			this.txtGastosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosNoviembre.Appearance = appearance16;
			this.txtGastosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosNoviembre.Enabled = false;
			this.txtGastosNoviembre.FormatString = "#,##0.00";
			this.txtGastosNoviembre.Location = new System.Drawing.Point(944, 405);
			this.txtGastosNoviembre.Name = "txtGastosNoviembre";
			this.txtGastosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosNoviembre.PromptChar = ' ';
			this.txtGastosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosNoviembre.TabIndex = 891;
			// 
			// txtGastosDiciembre
			// 
			this.txtGastosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosDiciembre.Appearance = appearance17;
			this.txtGastosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosDiciembre.Enabled = false;
			this.txtGastosDiciembre.FormatString = "#,##0.00";
			this.txtGastosDiciembre.Location = new System.Drawing.Point(1032, 405);
			this.txtGastosDiciembre.Name = "txtGastosDiciembre";
			this.txtGastosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosDiciembre.PromptChar = ' ';
			this.txtGastosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosDiciembre.TabIndex = 890;
			// 
			// txtIngresosDiciembre
			// 
			this.txtIngresosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosDiciembre.Appearance = appearance18;
			this.txtIngresosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosDiciembre.Enabled = false;
			this.txtIngresosDiciembre.FormatString = "#,##0.00";
			this.txtIngresosDiciembre.Location = new System.Drawing.Point(1032, 381);
			this.txtIngresosDiciembre.Name = "txtIngresosDiciembre";
			this.txtIngresosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosDiciembre.PromptChar = ' ';
			this.txtIngresosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosDiciembre.TabIndex = 889;
			// 
			// txtIngresosNoviembre
			// 
			this.txtIngresosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosNoviembre.Appearance = appearance19;
			this.txtIngresosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosNoviembre.Enabled = false;
			this.txtIngresosNoviembre.FormatString = "#,##0.00";
			this.txtIngresosNoviembre.Location = new System.Drawing.Point(944, 381);
			this.txtIngresosNoviembre.Name = "txtIngresosNoviembre";
			this.txtIngresosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosNoviembre.PromptChar = ' ';
			this.txtIngresosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosNoviembre.TabIndex = 888;
			// 
			// txtGastosSeptiembre
			// 
			this.txtGastosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosSeptiembre.Appearance = appearance20;
			this.txtGastosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosSeptiembre.Enabled = false;
			this.txtGastosSeptiembre.FormatString = "#,##0.00";
			this.txtGastosSeptiembre.Location = new System.Drawing.Point(768, 405);
			this.txtGastosSeptiembre.Name = "txtGastosSeptiembre";
			this.txtGastosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosSeptiembre.PromptChar = ' ';
			this.txtGastosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosSeptiembre.TabIndex = 887;
			// 
			// txtGastosOctubre
			// 
			this.txtGastosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosOctubre.Appearance = appearance21;
			this.txtGastosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosOctubre.Enabled = false;
			this.txtGastosOctubre.FormatString = "#,##0.00";
			this.txtGastosOctubre.Location = new System.Drawing.Point(856, 405);
			this.txtGastosOctubre.Name = "txtGastosOctubre";
			this.txtGastosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosOctubre.PromptChar = ' ';
			this.txtGastosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosOctubre.TabIndex = 886;
			// 
			// txtIngresosOctubre
			// 
			this.txtIngresosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosOctubre.Appearance = appearance22;
			this.txtIngresosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosOctubre.Enabled = false;
			this.txtIngresosOctubre.FormatString = "#,##0.00";
			this.txtIngresosOctubre.Location = new System.Drawing.Point(856, 381);
			this.txtIngresosOctubre.Name = "txtIngresosOctubre";
			this.txtIngresosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosOctubre.PromptChar = ' ';
			this.txtIngresosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosOctubre.TabIndex = 885;
			// 
			// txtIngresosSeptiembre
			// 
			this.txtIngresosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosSeptiembre.Appearance = appearance23;
			this.txtIngresosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosSeptiembre.Enabled = false;
			this.txtIngresosSeptiembre.FormatString = "#,##0.00";
			this.txtIngresosSeptiembre.Location = new System.Drawing.Point(768, 381);
			this.txtIngresosSeptiembre.Name = "txtIngresosSeptiembre";
			this.txtIngresosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosSeptiembre.PromptChar = ' ';
			this.txtIngresosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosSeptiembre.TabIndex = 884;
			// 
			// txtGastosJulio
			// 
			this.txtGastosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJulio.Appearance = appearance24;
			this.txtGastosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJulio.Enabled = false;
			this.txtGastosJulio.FormatString = "#,##0.00";
			this.txtGastosJulio.Location = new System.Drawing.Point(592, 405);
			this.txtGastosJulio.Name = "txtGastosJulio";
			this.txtGastosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJulio.PromptChar = ' ';
			this.txtGastosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJulio.TabIndex = 883;
			// 
			// txtGastosAgosto
			// 
			this.txtGastosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAgosto.Appearance = appearance25;
			this.txtGastosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAgosto.Enabled = false;
			this.txtGastosAgosto.FormatString = "#,##0.00";
			this.txtGastosAgosto.Location = new System.Drawing.Point(680, 405);
			this.txtGastosAgosto.Name = "txtGastosAgosto";
			this.txtGastosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAgosto.PromptChar = ' ';
			this.txtGastosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAgosto.TabIndex = 882;
			// 
			// txtIngresosAgosto
			// 
			this.txtIngresosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAgosto.Appearance = appearance26;
			this.txtIngresosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAgosto.Enabled = false;
			this.txtIngresosAgosto.FormatString = "#,##0.00";
			this.txtIngresosAgosto.Location = new System.Drawing.Point(680, 381);
			this.txtIngresosAgosto.Name = "txtIngresosAgosto";
			this.txtIngresosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAgosto.PromptChar = ' ';
			this.txtIngresosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAgosto.TabIndex = 881;
			// 
			// txtIngresosJulio
			// 
			this.txtIngresosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJulio.Appearance = appearance27;
			this.txtIngresosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJulio.Enabled = false;
			this.txtIngresosJulio.FormatString = "#,##0.00";
			this.txtIngresosJulio.Location = new System.Drawing.Point(592, 381);
			this.txtIngresosJulio.Name = "txtIngresosJulio";
			this.txtIngresosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJulio.PromptChar = ' ';
			this.txtIngresosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJulio.TabIndex = 880;
			// 
			// txtGastosMayo
			// 
			this.txtGastosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMayo.Appearance = appearance28;
			this.txtGastosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMayo.Enabled = false;
			this.txtGastosMayo.FormatString = "#,##0.00";
			this.txtGastosMayo.Location = new System.Drawing.Point(416, 405);
			this.txtGastosMayo.Name = "txtGastosMayo";
			this.txtGastosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMayo.PromptChar = ' ';
			this.txtGastosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMayo.TabIndex = 879;
			// 
			// txtGastosJunio
			// 
			this.txtGastosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJunio.Appearance = appearance29;
			this.txtGastosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJunio.Enabled = false;
			this.txtGastosJunio.FormatString = "#,##0.00";
			this.txtGastosJunio.Location = new System.Drawing.Point(504, 405);
			this.txtGastosJunio.Name = "txtGastosJunio";
			this.txtGastosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJunio.PromptChar = ' ';
			this.txtGastosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJunio.TabIndex = 878;
			// 
			// txtIngresosJunio
			// 
			this.txtIngresosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJunio.Appearance = appearance30;
			this.txtIngresosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJunio.Enabled = false;
			this.txtIngresosJunio.FormatString = "#,##0.00";
			this.txtIngresosJunio.Location = new System.Drawing.Point(504, 381);
			this.txtIngresosJunio.Name = "txtIngresosJunio";
			this.txtIngresosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJunio.PromptChar = ' ';
			this.txtIngresosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJunio.TabIndex = 877;
			// 
			// txtIngresosMayo
			// 
			this.txtIngresosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMayo.Appearance = appearance31;
			this.txtIngresosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMayo.Enabled = false;
			this.txtIngresosMayo.FormatString = "#,##0.00";
			this.txtIngresosMayo.Location = new System.Drawing.Point(416, 381);
			this.txtIngresosMayo.Name = "txtIngresosMayo";
			this.txtIngresosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMayo.PromptChar = ' ';
			this.txtIngresosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMayo.TabIndex = 876;
			// 
			// txtGastosMarzo
			// 
			this.txtGastosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMarzo.Appearance = appearance32;
			this.txtGastosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMarzo.Enabled = false;
			this.txtGastosMarzo.FormatString = "#,##0.00";
			this.txtGastosMarzo.Location = new System.Drawing.Point(240, 405);
			this.txtGastosMarzo.Name = "txtGastosMarzo";
			this.txtGastosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMarzo.PromptChar = ' ';
			this.txtGastosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMarzo.TabIndex = 875;
			// 
			// txtGastosAbril
			// 
			this.txtGastosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAbril.Appearance = appearance33;
			this.txtGastosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAbril.Enabled = false;
			this.txtGastosAbril.FormatString = "#,##0.00";
			this.txtGastosAbril.Location = new System.Drawing.Point(328, 405);
			this.txtGastosAbril.Name = "txtGastosAbril";
			this.txtGastosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAbril.PromptChar = ' ';
			this.txtGastosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAbril.TabIndex = 874;
			// 
			// txtIngresosAbril
			// 
			this.txtIngresosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAbril.Appearance = appearance34;
			this.txtIngresosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAbril.Enabled = false;
			this.txtIngresosAbril.FormatString = "#,##0.00";
			this.txtIngresosAbril.Location = new System.Drawing.Point(328, 381);
			this.txtIngresosAbril.Name = "txtIngresosAbril";
			this.txtIngresosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAbril.PromptChar = ' ';
			this.txtIngresosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAbril.TabIndex = 873;
			// 
			// txtIngresosMarzo
			// 
			this.txtIngresosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMarzo.Appearance = appearance35;
			this.txtIngresosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMarzo.Enabled = false;
			this.txtIngresosMarzo.FormatString = "#,##0.00";
			this.txtIngresosMarzo.Location = new System.Drawing.Point(240, 381);
			this.txtIngresosMarzo.Name = "txtIngresosMarzo";
			this.txtIngresosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMarzo.PromptChar = ' ';
			this.txtIngresosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMarzo.TabIndex = 872;
			// 
			// txtGastosEnero
			// 
			this.txtGastosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosEnero.Appearance = appearance36;
			this.txtGastosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosEnero.Enabled = false;
			this.txtGastosEnero.FormatString = "#,##0.00";
			this.txtGastosEnero.Location = new System.Drawing.Point(64, 405);
			this.txtGastosEnero.Name = "txtGastosEnero";
			this.txtGastosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosEnero.PromptChar = ' ';
			this.txtGastosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosEnero.TabIndex = 871;
			// 
			// txtGastosFebrero
			// 
			this.txtGastosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosFebrero.Appearance = appearance37;
			this.txtGastosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosFebrero.Enabled = false;
			this.txtGastosFebrero.FormatString = "#,##0.00";
			this.txtGastosFebrero.Location = new System.Drawing.Point(152, 405);
			this.txtGastosFebrero.Name = "txtGastosFebrero";
			this.txtGastosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFebrero.PromptChar = ' ';
			this.txtGastosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosFebrero.TabIndex = 870;
			// 
			// txtIngresosFebrero
			// 
			this.txtIngresosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosFebrero.Appearance = appearance38;
			this.txtIngresosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosFebrero.Enabled = false;
			this.txtIngresosFebrero.FormatString = "#,##0.00";
			this.txtIngresosFebrero.Location = new System.Drawing.Point(152, 381);
			this.txtIngresosFebrero.Name = "txtIngresosFebrero";
			this.txtIngresosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosFebrero.PromptChar = ' ';
			this.txtIngresosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosFebrero.TabIndex = 869;
			// 
			// txtIngresosEnero
			// 
			this.txtIngresosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosEnero.Appearance = appearance39;
			this.txtIngresosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosEnero.Enabled = false;
			this.txtIngresosEnero.FormatString = "#,##0.00";
			this.txtIngresosEnero.Location = new System.Drawing.Point(64, 381);
			this.txtIngresosEnero.Name = "txtIngresosEnero";
			this.txtIngresosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosEnero.PromptChar = ' ';
			this.txtIngresosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosEnero.TabIndex = 868;
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource4;
			appearance40.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance40;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Gasto";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance41;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 78;
			ultraGridColumn4.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance42;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 78;
			ultraGridColumn5.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance43;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 78;
			ultraGridColumn6.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance44;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 78;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance45;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 78;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance46;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 78;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance47;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 78;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance48;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 78;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance49;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 78;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance50;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 78;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance51;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 78;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance52;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 78;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance53;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 132;
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
																										 ultraGridColumn15});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance54;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance55;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance56;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance57;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance58;
			summarySettings5.DisplayFormat = "{0:#,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance59;
			summarySettings6.DisplayFormat = "{0:#,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance60;
			summarySettings7.DisplayFormat = "{0:#,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance61;
			summarySettings8.DisplayFormat = "{0:#,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance62;
			summarySettings9.DisplayFormat = "{0:#,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance63;
			summarySettings10.DisplayFormat = "{0:#,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance64;
			summarySettings11.DisplayFormat = "{0:#,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance65;
			summarySettings12.DisplayFormat = "{0:#,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance66;
			summarySettings13.DisplayFormat = "{0:#,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13});
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance67;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance68.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance68;
			appearance69.ForeColor = System.Drawing.Color.Black;
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance69;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance70.FontData.BoldAsString = "True";
			appearance70.FontData.Name = "Arial";
			appearance70.FontData.SizeInPoints = 8F;
			appearance70.ForeColor = System.Drawing.Color.White;
			appearance70.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance70;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance71.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance71.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance73;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 184);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1226, 168);
			this.uGridGastos.TabIndex = 866;
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
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
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn15});
			// 
			// uGridVentasPrecios
			// 
			this.uGridVentasPrecios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasPrecios.DataSource = this.ultraDataSource2;
			appearance74.BackColor = System.Drawing.Color.White;
			this.uGridVentasPrecios.DisplayLayout.Appearance = appearance74;
			this.uGridVentasPrecios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Format = "";
			ultraGridColumn17.Header.Caption = "BODEGA";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 132;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance75;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 82;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance76;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 82;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance77;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 82;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance78;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 82;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance79;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Width = 82;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance80;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Width = 82;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance81;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Width = 82;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance82;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 9;
			ultraGridColumn25.Width = 82;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance83;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 10;
			ultraGridColumn26.Width = 82;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance84;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 11;
			ultraGridColumn27.Width = 82;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance85;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 12;
			ultraGridColumn28.Width = 82;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance86;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 13;
			ultraGridColumn29.Width = 82;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance87;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Width = 89;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn30});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance88;
			summarySettings14.DisplayFormat = "{0:#,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance89;
			summarySettings15.DisplayFormat = "{0:#,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance90;
			summarySettings16.DisplayFormat = "{0:#,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance91;
			summarySettings17.DisplayFormat = "{0:#,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance92;
			summarySettings18.DisplayFormat = "{0:#,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance93;
			summarySettings19.DisplayFormat = "{0:#,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance94;
			summarySettings20.DisplayFormat = "{0:#,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance95;
			summarySettings21.DisplayFormat = "{0:#,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance96;
			summarySettings22.DisplayFormat = "{0:#,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance97;
			summarySettings23.DisplayFormat = "{0:#,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance98;
			summarySettings24.DisplayFormat = "{0:#,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance99;
			summarySettings25.DisplayFormat = "{0:#,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance100;
			summarySettings26.DisplayFormat = "{0:#,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridVentasPrecios.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance101.ForeColor = System.Drawing.Color.Black;
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.ActiveRowAppearance = appearance101;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance102.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.CardAreaAppearance = appearance102;
			appearance103.ForeColor = System.Drawing.Color.Black;
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.CellAppearance = appearance103;
			this.uGridVentasPrecios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance104.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance104.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance104.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance104.FontData.BoldAsString = "True";
			appearance104.FontData.Name = "Arial";
			appearance104.FontData.SizeInPoints = 8F;
			appearance104.ForeColor = System.Drawing.Color.White;
			appearance104.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.HeaderAppearance = appearance104;
			this.uGridVentasPrecios.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance105.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance105.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowAlternateAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowSelectorAppearance = appearance106;
			appearance107.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance107.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance107.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.SelectedRowAppearance = appearance107;
			this.uGridVentasPrecios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasPrecios.Location = new System.Drawing.Point(8, 40);
			this.uGridVentasPrecios.Name = "uGridVentasPrecios";
			this.uGridVentasPrecios.Size = new System.Drawing.Size(1226, 136);
			this.uGridVentasPrecios.TabIndex = 865;
			this.uGridVentasPrecios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasPrecios_KeyDown);
			this.uGridVentasPrecios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasPrecios_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
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
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn30});
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
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1152, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(82, 24);
			this.btnSalir.TabIndex = 926;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.button1_Click);
			// 
			// frmPresupuestoRealDentalis
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1240, 470);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.label17);
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
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.uGridVentasPrecios);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPresupuestoRealDentalis";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoRealDentalis_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoRealDentalis_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).EndInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPresupuestoRealDentalis_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmPresupuestoRealDentalis_Load(object sender, System.EventArgs e)
		{
//			this.txtAño.MinValue = 2020;
//			this.txtAño.MaxValue = DateTime.Today.Year + 1;
//			this.txtAño.Value = DateTime.Today.Year;

			string sSQL = string.Format("Exec [PresupuestoRealDentalisVenta] {0}", iYear);
			this.uGridVentasPrecios.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			this.SumatoriasAlConsultar(this.uGridVentasPrecios);

			string sSQLg = string.Format("Exec [ReportePeriodoGastoDentalis] {0}", iYear);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLg);

			this.SumatoriasAlConsultar(this.uGridGastos);
			this.Resultado();
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

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
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

		private void uGridVentasPrecios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

		private void uGridVentasPrecios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentasPrecios);
		}
		private void SumatoriasAlConsultar(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if ( dr.Cells["ENERO"].Value == System.DBNull.Value)(dr.Cells["ENERO"].Value) = 0;
				if ( dr.Cells["FEBRERO"].Value == System.DBNull.Value)(dr.Cells["FEBRERO"].Value) = 0;
				if ( dr.Cells["MARZO"].Value == System.DBNull.Value)(dr.Cells["MARZO"].Value) = 0;
				if ( dr.Cells["ABRIL"].Value == System.DBNull.Value)(dr.Cells["ABRIL"].Value) = 0;
				if ( dr.Cells["MAYO"].Value == System.DBNull.Value)(dr.Cells["MAYO"].Value) = 0;
				if ( dr.Cells["JUNIO"].Value == System.DBNull.Value)(dr.Cells["JUNIO"].Value) = 0;
				if ( dr.Cells["JULIO"].Value == System.DBNull.Value)(dr.Cells["JULIO"].Value) = 0;
				if ( dr.Cells["AGOSTO"].Value == System.DBNull.Value)(dr.Cells["AGOSTO"].Value) = 0;
				if ( dr.Cells["SEPTIEMBRE"].Value == System.DBNull.Value)(dr.Cells["SEPTIEMBRE"].Value) = 0;
				if ( dr.Cells["OCTUBRE"].Value == System.DBNull.Value)(dr.Cells["OCTUBRE"].Value) = 0;
				if ( dr.Cells["NOVIEMBRE"].Value == System.DBNull.Value)(dr.Cells["NOVIEMBRE"].Value) = 0;
				if ( dr.Cells["DICIEMBRE"].Value == System.DBNull.Value)(dr.Cells["DICIEMBRE"].Value) = 0;
	
			}

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

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGastos);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

	}
}
