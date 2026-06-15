using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoDentalis.
	/// </summary>
	public class frmPresupuestoDentalis : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasPrecios;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Button btnProyeccion;
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
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoDentalis()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoDentalis));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 2, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 3, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 4, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 5, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 6, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 7, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 8, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 9, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 10, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 11, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 14, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
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
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 2, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 3, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 4, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 5, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 6, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 7, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 8, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 9, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 10, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 11, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 14, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
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
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridVentasPrecios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label17 = new System.Windows.Forms.Label();
			this.btnProyeccion = new System.Windows.Forms.Button();
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
			this.btnNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
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
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(8, 8);
			this.txtAño.MinValue = 2017;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 741;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.Location = new System.Drawing.Point(112, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(30, 23);
			this.btnVer.TabIndex = 742;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridVentasPrecios
			// 
			this.uGridVentasPrecios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasPrecios.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridVentasPrecios.DisplayLayout.Appearance = appearance1;
			this.uGridVentasPrecios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Format = "";
			ultraGridColumn2.Header.Caption = "BODEGA";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 134;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 83;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 78;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 74;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 73;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 73;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 70;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 70;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance10;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 86;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 88;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance12;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 89;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance13;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 89;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 101;
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
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance17;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance18;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance19;
			summarySettings5.DisplayFormat = "{0:#,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance20;
			summarySettings6.DisplayFormat = "{0:#,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance21;
			summarySettings7.DisplayFormat = "{0:#,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance22;
			summarySettings8.DisplayFormat = "{0:#,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance23;
			summarySettings9.DisplayFormat = "{0:#,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance24;
			summarySettings10.DisplayFormat = "{0:#,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance25;
			summarySettings11.DisplayFormat = "{0:#,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance26;
			summarySettings12.DisplayFormat = "{0:#,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance27;
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
			this.uGridVentasPrecios.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.ActiveRowAppearance = appearance28;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance29.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.CardAreaAppearance = appearance29;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.CellAppearance = appearance30;
			this.uGridVentasPrecios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.HeaderAppearance = appearance31;
			appearance32.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance32.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowAlternateAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridVentasPrecios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasPrecios.Location = new System.Drawing.Point(8, 40);
			this.uGridVentasPrecios.Name = "uGridVentasPrecios";
			this.uGridVentasPrecios.Size = new System.Drawing.Size(1226, 112);
			this.uGridVentasPrecios.TabIndex = 743;
			this.uGridVentasPrecios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasPrecios_KeyDown);
			this.uGridVentasPrecios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVentasPrecios_AfterCellUpdate);
			this.uGridVentasPrecios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasPrecios_InitializeLayout);
			// 
			// ultraDataSource2
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
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource4;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance35;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Gasto";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 120;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance36;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 78;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance37;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 78;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance38;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 78;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance39;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 78;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance40;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 78;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance41;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 78;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance42;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 78;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance43;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 9;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 78;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance44;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 10;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 78;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance45;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 11;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 78;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance46;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 12;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 78;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance47;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 13;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 78;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance48;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Width = 132;
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
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance49;
			summarySettings14.DisplayFormat = "{0:#,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance50;
			summarySettings15.DisplayFormat = "{0:#,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance51;
			summarySettings16.DisplayFormat = "{0:#,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance52;
			summarySettings17.DisplayFormat = "{0:#,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance53;
			summarySettings18.DisplayFormat = "{0:#,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance54;
			summarySettings19.DisplayFormat = "{0:#,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance55;
			summarySettings20.DisplayFormat = "{0:#,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance56;
			summarySettings21.DisplayFormat = "{0:#,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance57;
			summarySettings22.DisplayFormat = "{0:#,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance58;
			summarySettings23.DisplayFormat = "{0:#,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance59;
			summarySettings24.DisplayFormat = "{0:#,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance60;
			summarySettings25.DisplayFormat = "{0:#,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance61;
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
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance62.ForeColor = System.Drawing.Color.Black;
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance62;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance63.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance63;
			appearance64.ForeColor = System.Drawing.Color.Black;
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance64;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance65.FontData.BoldAsString = "True";
			appearance65.FontData.Name = "Arial";
			appearance65.FontData.SizeInPoints = 8F;
			appearance65.ForeColor = System.Drawing.Color.White;
			appearance65.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance65;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance66.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance66.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance67;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance68;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 160);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1226, 168);
			this.uGridGastos.TabIndex = 746;
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_ClickCellButton);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource4
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
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
			// label17
			// 
			this.label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label17.Location = new System.Drawing.Point(16, 424);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(1240, 1);
			this.label17.TabIndex = 862;
			// 
			// btnProyeccion
			// 
			this.btnProyeccion.CausesValidation = false;
			this.btnProyeccion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnProyeccion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnProyeccion.Image = ((System.Drawing.Image)(resources.GetObject("btnProyeccion.Image")));
			this.btnProyeccion.Location = new System.Drawing.Point(8, 336);
			this.btnProyeccion.Name = "btnProyeccion";
			this.btnProyeccion.Size = new System.Drawing.Size(30, 23);
			this.btnProyeccion.TabIndex = 861;
			this.btnProyeccion.Click += new System.EventHandler(this.btnProyeccion_Click);
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Firebrick;
			this.label16.Location = new System.Drawing.Point(1120, 352);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(100, 16);
			this.label16.TabIndex = 860;
			this.label16.Text = "TOTAL";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(1032, 352);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(80, 16);
			this.label12.TabIndex = 859;
			this.label12.Text = "DICIEMBRE";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(944, 352);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 858;
			this.label13.Text = "NOVIEMBRE";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(856, 352);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(80, 16);
			this.label14.TabIndex = 857;
			this.label14.Text = "OCTUBRE";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(768, 352);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 856;
			this.label15.Text = "SEPTIEMBRE";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(680, 352);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 855;
			this.label8.Text = "AGOSTO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(592, 352);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 854;
			this.label9.Text = "JULIO";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(504, 352);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 853;
			this.label10.Text = "JUNIO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(416, 352);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 852;
			this.label11.Text = "MAYO";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(328, 352);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 851;
			this.label7.Text = "ABRIL";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(240, 352);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 850;
			this.label6.Text = "MARZO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(152, 352);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 849;
			this.label4.Text = "FEBRERO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(64, 352);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 848;
			this.label3.Text = "ENERO";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtUtilidadTotal
			// 
			this.txtUtilidadTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance69.FontData.BoldAsString = "True";
			appearance69.ForeColor = System.Drawing.Color.Firebrick;
			appearance69.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotal.Appearance = appearance69;
			this.txtUtilidadTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotal.Enabled = false;
			this.txtUtilidadTotal.FormatString = "#,##0.00";
			this.txtUtilidadTotal.Location = new System.Drawing.Point(1120, 432);
			this.txtUtilidadTotal.Name = "txtUtilidadTotal";
			this.txtUtilidadTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotal.PromptChar = ' ';
			this.txtUtilidadTotal.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotal.TabIndex = 847;
			// 
			// txtGastosTotal
			// 
			this.txtGastosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance70.FontData.BoldAsString = "True";
			appearance70.ForeColor = System.Drawing.Color.Firebrick;
			appearance70.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGastosTotal.Appearance = appearance70;
			this.txtGastosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosTotal.Enabled = false;
			this.txtGastosTotal.FormatString = "#,##0.00";
			this.txtGastosTotal.Location = new System.Drawing.Point(1120, 400);
			this.txtGastosTotal.Name = "txtGastosTotal";
			this.txtGastosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosTotal.PromptChar = ' ';
			this.txtGastosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtGastosTotal.TabIndex = 846;
			// 
			// txtIngresosTotal
			// 
			this.txtIngresosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance71.FontData.BoldAsString = "True";
			appearance71.ForeColor = System.Drawing.Color.Firebrick;
			appearance71.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtIngresosTotal.Appearance = appearance71;
			this.txtIngresosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosTotal.Enabled = false;
			this.txtIngresosTotal.FormatString = "#,##0.00";
			this.txtIngresosTotal.Location = new System.Drawing.Point(1120, 376);
			this.txtIngresosTotal.Name = "txtIngresosTotal";
			this.txtIngresosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosTotal.PromptChar = ' ';
			this.txtIngresosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtIngresosTotal.TabIndex = 845;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(24, 432);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 844;
			this.label2.Text = "(=) U";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(24, 400);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 843;
			this.label1.Text = "(-) G";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(24, 376);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 842;
			this.label5.Text = "(+) I";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUtilidadNoviembre
			// 
			this.txtUtilidadNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance72.ForeColor = System.Drawing.Color.Firebrick;
			appearance72.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadNoviembre.Appearance = appearance72;
			this.txtUtilidadNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadNoviembre.Enabled = false;
			this.txtUtilidadNoviembre.FormatString = "#,##0.00";
			this.txtUtilidadNoviembre.Location = new System.Drawing.Point(944, 432);
			this.txtUtilidadNoviembre.Name = "txtUtilidadNoviembre";
			this.txtUtilidadNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadNoviembre.PromptChar = ' ';
			this.txtUtilidadNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadNoviembre.TabIndex = 841;
			// 
			// txtUtilidadDiciembre
			// 
			this.txtUtilidadDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance73.ForeColor = System.Drawing.Color.Firebrick;
			appearance73.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadDiciembre.Appearance = appearance73;
			this.txtUtilidadDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadDiciembre.Enabled = false;
			this.txtUtilidadDiciembre.FormatString = "#,##0.00";
			this.txtUtilidadDiciembre.Location = new System.Drawing.Point(1032, 432);
			this.txtUtilidadDiciembre.Name = "txtUtilidadDiciembre";
			this.txtUtilidadDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadDiciembre.PromptChar = ' ';
			this.txtUtilidadDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadDiciembre.TabIndex = 840;
			// 
			// txtUtilidadSeptiembre
			// 
			this.txtUtilidadSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance74.ForeColor = System.Drawing.Color.Firebrick;
			appearance74.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadSeptiembre.Appearance = appearance74;
			this.txtUtilidadSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadSeptiembre.Enabled = false;
			this.txtUtilidadSeptiembre.FormatString = "#,##0.00";
			this.txtUtilidadSeptiembre.Location = new System.Drawing.Point(768, 432);
			this.txtUtilidadSeptiembre.Name = "txtUtilidadSeptiembre";
			this.txtUtilidadSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadSeptiembre.PromptChar = ' ';
			this.txtUtilidadSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadSeptiembre.TabIndex = 839;
			// 
			// txtUtilidadOctubre
			// 
			this.txtUtilidadOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance75.ForeColor = System.Drawing.Color.Firebrick;
			appearance75.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadOctubre.Appearance = appearance75;
			this.txtUtilidadOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadOctubre.Enabled = false;
			this.txtUtilidadOctubre.FormatString = "#,##0.00";
			this.txtUtilidadOctubre.Location = new System.Drawing.Point(856, 432);
			this.txtUtilidadOctubre.Name = "txtUtilidadOctubre";
			this.txtUtilidadOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadOctubre.PromptChar = ' ';
			this.txtUtilidadOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadOctubre.TabIndex = 838;
			// 
			// txtUtilidadJulio
			// 
			this.txtUtilidadJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance76.ForeColor = System.Drawing.Color.Firebrick;
			appearance76.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJulio.Appearance = appearance76;
			this.txtUtilidadJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJulio.Enabled = false;
			this.txtUtilidadJulio.FormatString = "#,##0.00";
			this.txtUtilidadJulio.Location = new System.Drawing.Point(592, 432);
			this.txtUtilidadJulio.Name = "txtUtilidadJulio";
			this.txtUtilidadJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJulio.PromptChar = ' ';
			this.txtUtilidadJulio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJulio.TabIndex = 837;
			// 
			// txtUtilidadAgosto
			// 
			this.txtUtilidadAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance77.ForeColor = System.Drawing.Color.Firebrick;
			appearance77.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAgosto.Appearance = appearance77;
			this.txtUtilidadAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAgosto.Enabled = false;
			this.txtUtilidadAgosto.FormatString = "#,##0.00";
			this.txtUtilidadAgosto.Location = new System.Drawing.Point(680, 432);
			this.txtUtilidadAgosto.Name = "txtUtilidadAgosto";
			this.txtUtilidadAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAgosto.PromptChar = ' ';
			this.txtUtilidadAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAgosto.TabIndex = 836;
			// 
			// txtUtilidadMayo
			// 
			this.txtUtilidadMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance78.ForeColor = System.Drawing.Color.Firebrick;
			appearance78.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMayo.Appearance = appearance78;
			this.txtUtilidadMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMayo.Enabled = false;
			this.txtUtilidadMayo.FormatString = "#,##0.00";
			this.txtUtilidadMayo.Location = new System.Drawing.Point(416, 432);
			this.txtUtilidadMayo.Name = "txtUtilidadMayo";
			this.txtUtilidadMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMayo.PromptChar = ' ';
			this.txtUtilidadMayo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMayo.TabIndex = 835;
			// 
			// txtUtilidadJunio
			// 
			this.txtUtilidadJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance79.ForeColor = System.Drawing.Color.Firebrick;
			appearance79.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJunio.Appearance = appearance79;
			this.txtUtilidadJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJunio.Enabled = false;
			this.txtUtilidadJunio.FormatString = "#,##0.00";
			this.txtUtilidadJunio.Location = new System.Drawing.Point(504, 432);
			this.txtUtilidadJunio.Name = "txtUtilidadJunio";
			this.txtUtilidadJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJunio.PromptChar = ' ';
			this.txtUtilidadJunio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJunio.TabIndex = 834;
			// 
			// txtUtilidadMarzo
			// 
			this.txtUtilidadMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance80.ForeColor = System.Drawing.Color.Firebrick;
			appearance80.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMarzo.Appearance = appearance80;
			this.txtUtilidadMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMarzo.Enabled = false;
			this.txtUtilidadMarzo.FormatString = "#,##0.00";
			this.txtUtilidadMarzo.Location = new System.Drawing.Point(240, 432);
			this.txtUtilidadMarzo.Name = "txtUtilidadMarzo";
			this.txtUtilidadMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMarzo.PromptChar = ' ';
			this.txtUtilidadMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMarzo.TabIndex = 833;
			// 
			// txtUtilidadAbril
			// 
			this.txtUtilidadAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance81.ForeColor = System.Drawing.Color.Firebrick;
			appearance81.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAbril.Appearance = appearance81;
			this.txtUtilidadAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAbril.Enabled = false;
			this.txtUtilidadAbril.FormatString = "#,##0.00";
			this.txtUtilidadAbril.Location = new System.Drawing.Point(328, 432);
			this.txtUtilidadAbril.Name = "txtUtilidadAbril";
			this.txtUtilidadAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAbril.PromptChar = ' ';
			this.txtUtilidadAbril.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAbril.TabIndex = 832;
			// 
			// txtUtilidadEnero
			// 
			this.txtUtilidadEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance82.ForeColor = System.Drawing.Color.Firebrick;
			appearance82.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadEnero.Appearance = appearance82;
			this.txtUtilidadEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadEnero.Enabled = false;
			this.txtUtilidadEnero.FormatString = "#,##0.00";
			this.txtUtilidadEnero.Location = new System.Drawing.Point(64, 432);
			this.txtUtilidadEnero.Name = "txtUtilidadEnero";
			this.txtUtilidadEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadEnero.PromptChar = ' ';
			this.txtUtilidadEnero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadEnero.TabIndex = 831;
			// 
			// txtUtilidadFebrero
			// 
			this.txtUtilidadFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance83.ForeColor = System.Drawing.Color.Firebrick;
			appearance83.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadFebrero.Appearance = appearance83;
			this.txtUtilidadFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadFebrero.Enabled = false;
			this.txtUtilidadFebrero.FormatString = "#,##0.00";
			this.txtUtilidadFebrero.Location = new System.Drawing.Point(152, 432);
			this.txtUtilidadFebrero.Name = "txtUtilidadFebrero";
			this.txtUtilidadFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadFebrero.PromptChar = ' ';
			this.txtUtilidadFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadFebrero.TabIndex = 830;
			// 
			// txtGastosNoviembre
			// 
			this.txtGastosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosNoviembre.Appearance = appearance84;
			this.txtGastosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosNoviembre.Enabled = false;
			this.txtGastosNoviembre.FormatString = "#,##0.00";
			this.txtGastosNoviembre.Location = new System.Drawing.Point(944, 400);
			this.txtGastosNoviembre.Name = "txtGastosNoviembre";
			this.txtGastosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosNoviembre.PromptChar = ' ';
			this.txtGastosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosNoviembre.TabIndex = 829;
			// 
			// txtGastosDiciembre
			// 
			this.txtGastosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosDiciembre.Appearance = appearance85;
			this.txtGastosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosDiciembre.Enabled = false;
			this.txtGastosDiciembre.FormatString = "#,##0.00";
			this.txtGastosDiciembre.Location = new System.Drawing.Point(1032, 400);
			this.txtGastosDiciembre.Name = "txtGastosDiciembre";
			this.txtGastosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosDiciembre.PromptChar = ' ';
			this.txtGastosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosDiciembre.TabIndex = 828;
			// 
			// txtIngresosDiciembre
			// 
			this.txtIngresosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosDiciembre.Appearance = appearance86;
			this.txtIngresosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosDiciembre.Enabled = false;
			this.txtIngresosDiciembre.FormatString = "#,##0.00";
			this.txtIngresosDiciembre.Location = new System.Drawing.Point(1032, 376);
			this.txtIngresosDiciembre.Name = "txtIngresosDiciembre";
			this.txtIngresosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosDiciembre.PromptChar = ' ';
			this.txtIngresosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosDiciembre.TabIndex = 827;
			// 
			// txtIngresosNoviembre
			// 
			this.txtIngresosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosNoviembre.Appearance = appearance87;
			this.txtIngresosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosNoviembre.Enabled = false;
			this.txtIngresosNoviembre.FormatString = "#,##0.00";
			this.txtIngresosNoviembre.Location = new System.Drawing.Point(944, 376);
			this.txtIngresosNoviembre.Name = "txtIngresosNoviembre";
			this.txtIngresosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosNoviembre.PromptChar = ' ';
			this.txtIngresosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosNoviembre.TabIndex = 826;
			// 
			// txtGastosSeptiembre
			// 
			this.txtGastosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosSeptiembre.Appearance = appearance88;
			this.txtGastosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosSeptiembre.Enabled = false;
			this.txtGastosSeptiembre.FormatString = "#,##0.00";
			this.txtGastosSeptiembre.Location = new System.Drawing.Point(768, 400);
			this.txtGastosSeptiembre.Name = "txtGastosSeptiembre";
			this.txtGastosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosSeptiembre.PromptChar = ' ';
			this.txtGastosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosSeptiembre.TabIndex = 825;
			// 
			// txtGastosOctubre
			// 
			this.txtGastosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance89.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosOctubre.Appearance = appearance89;
			this.txtGastosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosOctubre.Enabled = false;
			this.txtGastosOctubre.FormatString = "#,##0.00";
			this.txtGastosOctubre.Location = new System.Drawing.Point(856, 400);
			this.txtGastosOctubre.Name = "txtGastosOctubre";
			this.txtGastosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosOctubre.PromptChar = ' ';
			this.txtGastosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosOctubre.TabIndex = 824;
			// 
			// txtIngresosOctubre
			// 
			this.txtIngresosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosOctubre.Appearance = appearance90;
			this.txtIngresosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosOctubre.Enabled = false;
			this.txtIngresosOctubre.FormatString = "#,##0.00";
			this.txtIngresosOctubre.Location = new System.Drawing.Point(856, 376);
			this.txtIngresosOctubre.Name = "txtIngresosOctubre";
			this.txtIngresosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosOctubre.PromptChar = ' ';
			this.txtIngresosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosOctubre.TabIndex = 823;
			// 
			// txtIngresosSeptiembre
			// 
			this.txtIngresosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosSeptiembre.Appearance = appearance91;
			this.txtIngresosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosSeptiembre.Enabled = false;
			this.txtIngresosSeptiembre.FormatString = "#,##0.00";
			this.txtIngresosSeptiembre.Location = new System.Drawing.Point(768, 376);
			this.txtIngresosSeptiembre.Name = "txtIngresosSeptiembre";
			this.txtIngresosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosSeptiembre.PromptChar = ' ';
			this.txtIngresosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosSeptiembre.TabIndex = 822;
			// 
			// txtGastosJulio
			// 
			this.txtGastosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJulio.Appearance = appearance92;
			this.txtGastosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJulio.Enabled = false;
			this.txtGastosJulio.FormatString = "#,##0.00";
			this.txtGastosJulio.Location = new System.Drawing.Point(592, 400);
			this.txtGastosJulio.Name = "txtGastosJulio";
			this.txtGastosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJulio.PromptChar = ' ';
			this.txtGastosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJulio.TabIndex = 821;
			// 
			// txtGastosAgosto
			// 
			this.txtGastosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAgosto.Appearance = appearance93;
			this.txtGastosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAgosto.Enabled = false;
			this.txtGastosAgosto.FormatString = "#,##0.00";
			this.txtGastosAgosto.Location = new System.Drawing.Point(680, 400);
			this.txtGastosAgosto.Name = "txtGastosAgosto";
			this.txtGastosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAgosto.PromptChar = ' ';
			this.txtGastosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAgosto.TabIndex = 820;
			// 
			// txtIngresosAgosto
			// 
			this.txtIngresosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAgosto.Appearance = appearance94;
			this.txtIngresosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAgosto.Enabled = false;
			this.txtIngresosAgosto.FormatString = "#,##0.00";
			this.txtIngresosAgosto.Location = new System.Drawing.Point(680, 376);
			this.txtIngresosAgosto.Name = "txtIngresosAgosto";
			this.txtIngresosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAgosto.PromptChar = ' ';
			this.txtIngresosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAgosto.TabIndex = 819;
			// 
			// txtIngresosJulio
			// 
			this.txtIngresosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJulio.Appearance = appearance95;
			this.txtIngresosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJulio.Enabled = false;
			this.txtIngresosJulio.FormatString = "#,##0.00";
			this.txtIngresosJulio.Location = new System.Drawing.Point(592, 376);
			this.txtIngresosJulio.Name = "txtIngresosJulio";
			this.txtIngresosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJulio.PromptChar = ' ';
			this.txtIngresosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJulio.TabIndex = 818;
			// 
			// txtGastosMayo
			// 
			this.txtGastosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMayo.Appearance = appearance96;
			this.txtGastosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMayo.Enabled = false;
			this.txtGastosMayo.FormatString = "#,##0.00";
			this.txtGastosMayo.Location = new System.Drawing.Point(416, 400);
			this.txtGastosMayo.Name = "txtGastosMayo";
			this.txtGastosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMayo.PromptChar = ' ';
			this.txtGastosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMayo.TabIndex = 817;
			// 
			// txtGastosJunio
			// 
			this.txtGastosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJunio.Appearance = appearance97;
			this.txtGastosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJunio.Enabled = false;
			this.txtGastosJunio.FormatString = "#,##0.00";
			this.txtGastosJunio.Location = new System.Drawing.Point(504, 400);
			this.txtGastosJunio.Name = "txtGastosJunio";
			this.txtGastosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJunio.PromptChar = ' ';
			this.txtGastosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJunio.TabIndex = 816;
			// 
			// txtIngresosJunio
			// 
			this.txtIngresosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJunio.Appearance = appearance98;
			this.txtIngresosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJunio.Enabled = false;
			this.txtIngresosJunio.FormatString = "#,##0.00";
			this.txtIngresosJunio.Location = new System.Drawing.Point(504, 376);
			this.txtIngresosJunio.Name = "txtIngresosJunio";
			this.txtIngresosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJunio.PromptChar = ' ';
			this.txtIngresosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJunio.TabIndex = 815;
			// 
			// txtIngresosMayo
			// 
			this.txtIngresosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMayo.Appearance = appearance99;
			this.txtIngresosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMayo.Enabled = false;
			this.txtIngresosMayo.FormatString = "#,##0.00";
			this.txtIngresosMayo.Location = new System.Drawing.Point(416, 376);
			this.txtIngresosMayo.Name = "txtIngresosMayo";
			this.txtIngresosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMayo.PromptChar = ' ';
			this.txtIngresosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMayo.TabIndex = 814;
			// 
			// txtGastosMarzo
			// 
			this.txtGastosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMarzo.Appearance = appearance100;
			this.txtGastosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMarzo.Enabled = false;
			this.txtGastosMarzo.FormatString = "#,##0.00";
			this.txtGastosMarzo.Location = new System.Drawing.Point(240, 400);
			this.txtGastosMarzo.Name = "txtGastosMarzo";
			this.txtGastosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMarzo.PromptChar = ' ';
			this.txtGastosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMarzo.TabIndex = 813;
			// 
			// txtGastosAbril
			// 
			this.txtGastosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAbril.Appearance = appearance101;
			this.txtGastosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAbril.Enabled = false;
			this.txtGastosAbril.FormatString = "#,##0.00";
			this.txtGastosAbril.Location = new System.Drawing.Point(328, 400);
			this.txtGastosAbril.Name = "txtGastosAbril";
			this.txtGastosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAbril.PromptChar = ' ';
			this.txtGastosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAbril.TabIndex = 812;
			// 
			// txtIngresosAbril
			// 
			this.txtIngresosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAbril.Appearance = appearance102;
			this.txtIngresosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAbril.Enabled = false;
			this.txtIngresosAbril.FormatString = "#,##0.00";
			this.txtIngresosAbril.Location = new System.Drawing.Point(328, 376);
			this.txtIngresosAbril.Name = "txtIngresosAbril";
			this.txtIngresosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAbril.PromptChar = ' ';
			this.txtIngresosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAbril.TabIndex = 811;
			// 
			// txtIngresosMarzo
			// 
			this.txtIngresosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMarzo.Appearance = appearance103;
			this.txtIngresosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMarzo.Enabled = false;
			this.txtIngresosMarzo.FormatString = "#,##0.00";
			this.txtIngresosMarzo.Location = new System.Drawing.Point(240, 376);
			this.txtIngresosMarzo.Name = "txtIngresosMarzo";
			this.txtIngresosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMarzo.PromptChar = ' ';
			this.txtIngresosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMarzo.TabIndex = 810;
			// 
			// txtGastosEnero
			// 
			this.txtGastosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosEnero.Appearance = appearance104;
			this.txtGastosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosEnero.Enabled = false;
			this.txtGastosEnero.FormatString = "#,##0.00";
			this.txtGastosEnero.Location = new System.Drawing.Point(64, 400);
			this.txtGastosEnero.Name = "txtGastosEnero";
			this.txtGastosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosEnero.PromptChar = ' ';
			this.txtGastosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosEnero.TabIndex = 809;
			// 
			// txtGastosFebrero
			// 
			this.txtGastosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance105.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosFebrero.Appearance = appearance105;
			this.txtGastosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosFebrero.Enabled = false;
			this.txtGastosFebrero.FormatString = "#,##0.00";
			this.txtGastosFebrero.Location = new System.Drawing.Point(152, 400);
			this.txtGastosFebrero.Name = "txtGastosFebrero";
			this.txtGastosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFebrero.PromptChar = ' ';
			this.txtGastosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosFebrero.TabIndex = 808;
			// 
			// txtIngresosFebrero
			// 
			this.txtIngresosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance106.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosFebrero.Appearance = appearance106;
			this.txtIngresosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosFebrero.Enabled = false;
			this.txtIngresosFebrero.FormatString = "#,##0.00";
			this.txtIngresosFebrero.Location = new System.Drawing.Point(152, 376);
			this.txtIngresosFebrero.Name = "txtIngresosFebrero";
			this.txtIngresosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosFebrero.PromptChar = ' ';
			this.txtIngresosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosFebrero.TabIndex = 807;
			// 
			// txtIngresosEnero
			// 
			this.txtIngresosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance107.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosEnero.Appearance = appearance107;
			this.txtIngresosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosEnero.Enabled = false;
			this.txtIngresosEnero.FormatString = "#,##0.00";
			this.txtIngresosEnero.Location = new System.Drawing.Point(64, 376);
			this.txtIngresosEnero.Name = "txtIngresosEnero";
			this.txtIngresosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosEnero.PromptChar = ' ';
			this.txtIngresosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosEnero.TabIndex = 806;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.Location = new System.Drawing.Point(8, 368);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(30, 23);
			this.btnNuevo.TabIndex = 805;
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			// 
			// frmPresupuestoDentalis
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1240, 470);
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
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.uGridVentasPrecios);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoDentalis";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Presupuesto Anual";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoDentalis_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoDentalis_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
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
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPresupuestoDentalis_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void uGridVentasPrecios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores ((int)e.Cell.Row.Cells["Bodega"].Value, "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}
		private void ActualizaValores (int sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec [ActualizaValoresPresupuestoAnual] {0}, {1}, '{2}', {3}", 
				Convert.ToInt32(this.txtAño.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
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

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtAño.Value == 0 || this.txtAño.Value  == System.DBNull.Value )
			{
				MessageBox.Show("Seleccione el Periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAño.Focus();
				return;
			}

			string sSQL = string.Format("Exec [consultaPresupuestoAnual] {0}", (int)this.txtAño.Value);
			this.uGridVentasPrecios.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			if ((int)this.uGridVentasPrecios.Rows.Count == 0)
			{
				string sSQLp = string.Format("Exec [GeneraPeriodoPResupuestoAnualGasto] {0}", 
					Convert.ToInt32(this.txtAño.Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLp);

				string sSQLa = string.Format("Exec [consultaPresupuestoAnual] {0}", (int)this.txtAño.Value);
				this.uGridVentasPrecios.DataSource = FuncionesProcedimientos.dtGeneral(sSQLa);
			}

			this.SumatoriasAlConsultar(this.uGridVentasPrecios);

			string sSQLg = string.Format("Exec [ReportePeriodoGasto] {0}", (int)this.txtAño.Value);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLg);

			this.SumatoriasAlConsultar(this.uGridGastos);
			this.Resultado();
			
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmPresupuestoDentalis_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2602");

			if (!Funcion.Permiso("1480", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Presupuesto DENTALIS ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.txtAño.MinValue = 2020;
			this.txtAño.MaxValue = DateTime.Today.Year + 1;
			this.txtAño.Value = DateTime.Today.Year;
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

		private void btnProyeccion_Click(object sender, System.EventArgs e)
		{
			using (frmPresupuestoRealDentalis CCJ = new frmPresupuestoRealDentalis (  (int)this.txtAño.Value ))
			{
				if (DialogResult.OK == CCJ.ShowDialog())
				{
						
				}
			}	
		}

		private void uGridGastos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (e.Cell.Column.ToString() == sColumna)
			{
				using (frmPresupuestoDentalisGastoDetalle CCJ = new frmPresupuestoDentalisGastoDetalle ( (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, 
								 sColumna, (int) this.txtAño.Value, (string)e.Cell.Row.Cells["Grupo"].Value))
				{
					if (DialogResult.OK == CCJ.ShowDialog())
					{
						
					}
				}
			}
		}
	}
}
