using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteGastos.
	/// </summary>
	public class frmReporteGastos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta2;
		private System.Windows.Forms.Button btnVer2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblContador2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.Button btnAcciones;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem mnuAcreditacionTransferencias;
		private System.Windows.Forms.MenuItem mnuAcreditacionCheques;
		private System.Windows.Forms.MenuItem menuItem1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridPorcentaje;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugtMEs;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteGastos()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteGastos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEnero", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFebrero", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMarzo", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAbril", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMayo", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJunio", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJulio", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAgosto", 7);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeptiembre", 8);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOctubre", 9);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNoviembre", 10);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDiciembre", 11);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEnero", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFebrero", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMarzo", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAbril", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMayo", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJunio", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJulio", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAgosto", 7);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeptiembre", 8);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOctubre", 9);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNoviembre", 10);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDiciembre", 11);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnEnero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnFebrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnMarzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnAbril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnMayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnJunio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnJulio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnAgosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnSeptiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnOctubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnNoviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnDiciembre");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEnero", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFebrero", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMarzo", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAbril", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMayo", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJunio", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJulio", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAgosto", 7);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeptiembre", 8);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOctubre", 9);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNoviembre", 10);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDiciembre", 11);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEnero", 0);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 2, true);
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGastos2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtDesde2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnVer2 = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblContador2 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.btnAcciones = new System.Windows.Forms.Button();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.mnuAcreditacionTransferencias = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionCheques = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ugridPorcentaje = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ugtMEs = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugtMEs)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(192, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 745;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 744;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 6);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 742;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 6);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 743;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.Caption = "Año";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Gasto";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 97;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 81;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 81;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 78;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 78;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 78;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Width = 78;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance10;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Width = 78;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 16;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance12;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 18;
			ultraGridColumn12.Width = 82;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 20;
			ultraGridColumn13.Width = 82;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 22;
			ultraGridColumn14.Width = 82;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance15;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 24;
			ultraGridColumn15.Width = 82;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance16;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 26;
			ultraGridColumn16.Width = 82;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 20;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 20;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 20;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 20;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 20;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 20;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn23.Header.Caption = "...";
			ultraGridColumn23.Header.VisiblePosition = 17;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 20;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn24.Header.Caption = "...";
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 20;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn25.Header.Caption = "...";
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 20;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn26.Header.Caption = "...";
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 20;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 25;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 20;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 20;
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
			ultraGridBand1.Header.Caption = "PERIODO A";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance17;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance18;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance19;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance20;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance21;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance22;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance23;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance24;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance25;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance26;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance27;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance28;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance29;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
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
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance30;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance32;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance33.FontData.BoldAsString = "True";
			appearance33.FontData.Name = "Arial";
			appearance33.FontData.SizeInPoints = 8F;
			appearance33.ForeColor = System.Drawing.Color.White;
			appearance33.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance33;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance34.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance34.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 48);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1215, 136);
			this.uGridGastos.TabIndex = 943;
			this.uGridGastos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_ClickCellButton);
			// 
			// ultraDataSource1
			// 
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
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// uGridGastos2
			// 
			this.uGridGastos2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos2.DataSource = this.ultraDataSource2;
			appearance37.BackColor = System.Drawing.Color.White;
			this.uGridGastos2.DisplayLayout.Appearance = appearance37;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Header.Caption = "Año";
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 137;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Gasto";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Width = 97;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance38;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Width = 81;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance39;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Width = 81;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance40;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Width = 78;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance41;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 8;
			ultraGridColumn35.Width = 78;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance42;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 10;
			ultraGridColumn36.Width = 78;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance43;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 12;
			ultraGridColumn37.Width = 78;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance44;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 14;
			ultraGridColumn38.Width = 78;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance45;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 16;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance46;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 18;
			ultraGridColumn40.Width = 82;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance47;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 20;
			ultraGridColumn41.Width = 82;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance48;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 22;
			ultraGridColumn42.Width = 82;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance49;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 24;
			ultraGridColumn43.Width = 82;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance50;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 26;
			ultraGridColumn44.Width = 82;
			ultraGridColumn45.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn45.Header.Caption = "...";
			ultraGridColumn45.Header.VisiblePosition = 5;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 20;
			ultraGridColumn46.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn46.Header.Caption = "...";
			ultraGridColumn46.Header.VisiblePosition = 7;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn46.Width = 20;
			ultraGridColumn47.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn47.Header.Caption = "...";
			ultraGridColumn47.Header.VisiblePosition = 9;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn47.Width = 20;
			ultraGridColumn48.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn48.Header.Caption = "...";
			ultraGridColumn48.Header.VisiblePosition = 11;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn48.Width = 20;
			ultraGridColumn49.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn49.Header.Caption = "...";
			ultraGridColumn49.Header.VisiblePosition = 13;
			ultraGridColumn49.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn49.Width = 20;
			ultraGridColumn50.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn50.Header.Caption = "...";
			ultraGridColumn50.Header.VisiblePosition = 15;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn50.Width = 20;
			ultraGridColumn51.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn51.Header.Caption = "...";
			ultraGridColumn51.Header.VisiblePosition = 17;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn51.Width = 20;
			ultraGridColumn52.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn52.Header.Caption = "...";
			ultraGridColumn52.Header.VisiblePosition = 19;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn52.Width = 20;
			ultraGridColumn53.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn53.Header.Caption = "...";
			ultraGridColumn53.Header.VisiblePosition = 21;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn53.Width = 20;
			ultraGridColumn54.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn54.Header.Caption = "...";
			ultraGridColumn54.Header.VisiblePosition = 23;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn54.Width = 20;
			ultraGridColumn55.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn55.Header.Caption = "...";
			ultraGridColumn55.Header.VisiblePosition = 25;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn55.Width = 20;
			ultraGridColumn56.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn56.Header.Caption = "...";
			ultraGridColumn56.Header.VisiblePosition = 27;
			ultraGridColumn56.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn56.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
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
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56});
			ultraGridBand2.Header.Caption = "PERIODO B";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance51;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance52;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance53;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance54;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance55;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance56;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance57;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance58;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance59;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance60;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance61;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance62;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance63;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
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
			this.uGridGastos2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance64.ForeColor = System.Drawing.Color.Black;
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos2.DisplayLayout.Override.ActiveRowAppearance = appearance64;
			this.uGridGastos2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance65.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos2.DisplayLayout.Override.CardAreaAppearance = appearance65;
			appearance66.ForeColor = System.Drawing.Color.Black;
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos2.DisplayLayout.Override.CellAppearance = appearance66;
			this.uGridGastos2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance67.FontData.BoldAsString = "True";
			appearance67.FontData.Name = "Arial";
			appearance67.FontData.SizeInPoints = 8F;
			appearance67.ForeColor = System.Drawing.Color.White;
			appearance67.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos2.DisplayLayout.Override.HeaderAppearance = appearance67;
			this.uGridGastos2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance68.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance68.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos2.DisplayLayout.Override.RowAlternateAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos2.DisplayLayout.Override.RowSelectorAppearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos2.DisplayLayout.Override.SelectedRowAppearance = appearance70;
			this.uGridGastos2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos2.Location = new System.Drawing.Point(8, 232);
			this.uGridGastos2.Name = "uGridGastos2";
			this.uGridGastos2.Size = new System.Drawing.Size(1215, 136);
			this.uGridGastos2.TabIndex = 944;
			this.uGridGastos2.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos2_ClickCellButton);
			// 
			// ultraDataSource2
			// 
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
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1240, 8);
			this.groupBox1.TabIndex = 945;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 224);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1240, 8);
			this.groupBox2.TabIndex = 950;
			this.groupBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(192, 200);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 949;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 200);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 948;
			this.label2.Text = "Desde";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde2
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde2.Appearance = appearance71;
			this.dtDesde2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtDesde2.DateButtons.Add(dateButton3);
			this.dtDesde2.Format = "dd/MM/yyyy";
			this.dtDesde2.Location = new System.Drawing.Point(56, 198);
			this.dtDesde2.Name = "dtDesde2";
			this.dtDesde2.NonAutoSizeHeight = 23;
			this.dtDesde2.Size = new System.Drawing.Size(112, 21);
			this.dtDesde2.SpinButtonsVisible = true;
			this.dtDesde2.TabIndex = 946;
			this.dtDesde2.Value = ((object)(resources.GetObject("dtDesde2.Value")));
			// 
			// dtHasta2
			// 
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta2.Appearance = appearance72;
			this.dtHasta2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtHasta2.DateButtons.Add(dateButton4);
			this.dtHasta2.Format = "dd/MM/yyyy";
			this.dtHasta2.Location = new System.Drawing.Point(248, 198);
			this.dtHasta2.Name = "dtHasta2";
			this.dtHasta2.NonAutoSizeHeight = 23;
			this.dtHasta2.Size = new System.Drawing.Size(112, 21);
			this.dtHasta2.SpinButtonsVisible = true;
			this.dtHasta2.TabIndex = 947;
			this.dtHasta2.Value = ((object)(resources.GetObject("dtHasta2.Value")));
			this.dtHasta2.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.ultraCalendarCombo2_BeforeDropDown);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(376, 4);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 951;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnVer2
			// 
			this.btnVer2.Image = ((System.Drawing.Image)(resources.GetObject("btnVer2.Image")));
			this.btnVer2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer2.Location = new System.Drawing.Point(376, 196);
			this.btnVer2.Name = "btnVer2";
			this.btnVer2.Size = new System.Drawing.Size(75, 24);
			this.btnVer2.TabIndex = 952;
			this.btnVer2.Text = "Ver";
			this.btnVer2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer2.Click += new System.EventHandler(this.btnVer2_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-4, 184);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1240, 8);
			this.groupBox3.TabIndex = 953;
			this.groupBox3.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Location = new System.Drawing.Point(0, 368);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1240, 8);
			this.groupBox4.TabIndex = 956;
			this.groupBox4.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(640, 6);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 957;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador2
			// 
			this.lblContador2.AutoSize = true;
			this.lblContador2.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador2.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador2.Location = new System.Drawing.Point(472, 198);
			this.lblContador2.Name = "lblContador2";
			this.lblContador2.Size = new System.Drawing.Size(0, 20);
			this.lblContador2.TabIndex = 958;
			this.lblContador2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn48.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
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
																																 ultraDataColumn60});
			// 
			// btnCalcular
			// 
			this.btnCalcular.Image = ((System.Drawing.Image)(resources.GetObject("btnCalcular.Image")));
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.Location = new System.Drawing.Point(1056, 4);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(75, 24);
			this.btnCalcular.TabIndex = 960;
			this.btnCalcular.Text = "Calcular";
			this.btnCalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btnAcciones
			// 
			this.btnAcciones.Enabled = false;
			this.btnAcciones.Image = ((System.Drawing.Image)(resources.GetObject("btnAcciones.Image")));
			this.btnAcciones.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcciones.Location = new System.Drawing.Point(1144, 4);
			this.btnAcciones.Name = "btnAcciones";
			this.btnAcciones.Size = new System.Drawing.Size(75, 24);
			this.btnAcciones.TabIndex = 961;
			this.btnAcciones.Text = "Acciones";
			this.btnAcciones.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcciones.Click += new System.EventHandler(this.btnAcciones_Click);
			// 
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAcreditacionTransferencias,
																																										this.mnuAcreditacionCheques,
																																										this.menuItem1});
			// 
			// mnuAcreditacionTransferencias
			// 
			this.mnuAcreditacionTransferencias.Index = 0;
			this.mnuAcreditacionTransferencias.Text = "Periodo A";
			this.mnuAcreditacionTransferencias.Click += new System.EventHandler(this.mnuAcreditacionTransferencias_Click);
			// 
			// mnuAcreditacionCheques
			// 
			this.mnuAcreditacionCheques.Index = 1;
			this.mnuAcreditacionCheques.Text = "Periodo B";
			this.mnuAcreditacionCheques.Click += new System.EventHandler(this.mnuAcreditacionCheques_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Porcentajes";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
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
			// ugridPorcentaje
			// 
			this.ugridPorcentaje.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridPorcentaje.DataSource = this.ultraDataSource2;
			appearance73.BackColor = System.Drawing.Color.White;
			this.ugridPorcentaje.DisplayLayout.Appearance = appearance73;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn57.Header.Caption = "Año";
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 137;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn59.Header.Caption = "Gasto";
			ultraGridColumn59.Header.VisiblePosition = 2;
			ultraGridColumn59.Width = 97;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance74;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 3;
			ultraGridColumn60.Width = 81;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance75;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn61.Width = 81;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance76;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 6;
			ultraGridColumn62.Width = 78;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance77;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 8;
			ultraGridColumn63.Width = 78;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance78;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 10;
			ultraGridColumn64.Width = 78;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance79;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 12;
			ultraGridColumn65.Width = 78;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance80;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 14;
			ultraGridColumn66.Width = 78;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance81;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 16;
			ultraGridColumn67.Width = 80;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance82;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 18;
			ultraGridColumn68.Width = 82;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance83;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 20;
			ultraGridColumn69.Width = 82;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance84;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 22;
			ultraGridColumn70.Width = 82;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance85;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 24;
			ultraGridColumn71.Width = 82;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance86;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 26;
			ultraGridColumn72.Width = 82;
			ultraGridColumn73.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn73.Header.Caption = "...";
			ultraGridColumn73.Header.VisiblePosition = 5;
			ultraGridColumn73.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn73.Width = 20;
			ultraGridColumn74.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn74.Header.Caption = "...";
			ultraGridColumn74.Header.VisiblePosition = 7;
			ultraGridColumn74.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn74.Width = 20;
			ultraGridColumn75.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn75.Header.Caption = "...";
			ultraGridColumn75.Header.VisiblePosition = 9;
			ultraGridColumn75.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn75.Width = 20;
			ultraGridColumn76.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn76.Header.Caption = "...";
			ultraGridColumn76.Header.VisiblePosition = 11;
			ultraGridColumn76.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn76.Width = 20;
			ultraGridColumn77.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn77.Header.Caption = "...";
			ultraGridColumn77.Header.VisiblePosition = 13;
			ultraGridColumn77.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn77.Width = 20;
			ultraGridColumn78.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn78.Header.Caption = "...";
			ultraGridColumn78.Header.VisiblePosition = 15;
			ultraGridColumn78.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn78.Width = 20;
			ultraGridColumn79.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn79.Header.Caption = "...";
			ultraGridColumn79.Header.VisiblePosition = 17;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn79.Width = 20;
			ultraGridColumn80.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn80.Header.Caption = "...";
			ultraGridColumn80.Header.VisiblePosition = 19;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn80.Width = 20;
			ultraGridColumn81.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn81.Header.Caption = "...";
			ultraGridColumn81.Header.VisiblePosition = 21;
			ultraGridColumn81.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn81.Width = 20;
			ultraGridColumn82.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn82.Header.Caption = "...";
			ultraGridColumn82.Header.VisiblePosition = 23;
			ultraGridColumn82.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn82.Width = 20;
			ultraGridColumn83.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn83.Header.Caption = "...";
			ultraGridColumn83.Header.VisiblePosition = 25;
			ultraGridColumn83.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn83.Width = 20;
			ultraGridColumn84.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn84.Header.Caption = "...";
			ultraGridColumn84.Header.VisiblePosition = 27;
			ultraGridColumn84.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn84.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
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
			ultraGridBand3.Header.Caption = "PERIODO B";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance87;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance88;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance89;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance90;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance91;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance92;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance93;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance94;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance95;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance96;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance97;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance98;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance99;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings39});
			ultraGridBand3.SummaryFooterCaption = "TOTALES";
			this.ugridPorcentaje.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance100.ForeColor = System.Drawing.Color.Black;
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.ActiveRowAppearance = appearance100;
			this.ugridPorcentaje.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridPorcentaje.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridPorcentaje.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance101.BackColor = System.Drawing.Color.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.CardAreaAppearance = appearance101;
			appearance102.ForeColor = System.Drawing.Color.Black;
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.CellAppearance = appearance102;
			this.ugridPorcentaje.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance103.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance103.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance103.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance103.FontData.BoldAsString = "True";
			appearance103.FontData.Name = "Arial";
			appearance103.FontData.SizeInPoints = 8F;
			appearance103.ForeColor = System.Drawing.Color.White;
			appearance103.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderAppearance = appearance103;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance104.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance104.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance104.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowAlternateAppearance = appearance104;
			appearance105.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance105.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowSelectorAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.SelectedRowAppearance = appearance106;
			this.ugridPorcentaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridPorcentaje.Location = new System.Drawing.Point(8, 384);
			this.ugridPorcentaje.Name = "ugridPorcentaje";
			this.ugridPorcentaje.Size = new System.Drawing.Size(1215, 216);
			this.ugridPorcentaje.TabIndex = 962;
			this.ugridPorcentaje.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugridPorcentaje_ClickCellButton_1);
			// 
			// optEstado
			// 
			appearance107.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance107;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.ItemAppearance = appearance108;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Periodo";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Mes";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(480, 3);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(136, 26);
			this.optEstado.TabIndex = 963;
			this.optEstado.Text = "Periodo";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// ugtMEs
			// 
			this.ugtMEs.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugtMEs.DataSource = this.ultraDataSource4;
			appearance109.BackColor = System.Drawing.Color.White;
			this.ugtMEs.DisplayLayout.Appearance = appearance109;
			ultraGridColumn85.Header.VisiblePosition = 0;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn86.Header.VisiblePosition = 1;
			ultraGridColumn86.Width = 180;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance110;
			ultraGridColumn87.Format = "#,##0.00";
			ultraGridColumn87.Header.VisiblePosition = 2;
			ultraGridColumn87.Width = 100;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance111;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellButtonAppearance = appearance112;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.VisiblePosition = 3;
			ultraGridColumn88.Width = 100;
			ultraGridColumn89.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn89.Header.Caption = "...";
			ultraGridColumn89.Header.VisiblePosition = 4;
			ultraGridColumn89.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn89.Width = 20;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89});
			ultraGridBand4.Header.Caption = "PERIODO B";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance113;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings40});
			ultraGridBand4.SummaryFooterCaption = "TOTALES";
			this.ugtMEs.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance114.ForeColor = System.Drawing.Color.Black;
			appearance114.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugtMEs.DisplayLayout.Override.ActiveRowAppearance = appearance114;
			this.ugtMEs.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugtMEs.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugtMEs.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance115.BackColor = System.Drawing.Color.Transparent;
			this.ugtMEs.DisplayLayout.Override.CardAreaAppearance = appearance115;
			appearance116.ForeColor = System.Drawing.Color.Black;
			appearance116.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugtMEs.DisplayLayout.Override.CellAppearance = appearance116;
			this.ugtMEs.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance117.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance117.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance117.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance117.FontData.BoldAsString = "True";
			appearance117.FontData.Name = "Arial";
			appearance117.FontData.SizeInPoints = 8F;
			appearance117.ForeColor = System.Drawing.Color.White;
			appearance117.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugtMEs.DisplayLayout.Override.HeaderAppearance = appearance117;
			this.ugtMEs.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance118.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance118.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance118.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugtMEs.DisplayLayout.Override.RowAlternateAppearance = appearance118;
			appearance119.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance119.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance119.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugtMEs.DisplayLayout.Override.RowSelectorAppearance = appearance119;
			appearance120.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance120.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance120.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugtMEs.DisplayLayout.Override.SelectedRowAppearance = appearance120;
			this.ugtMEs.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugtMEs.Location = new System.Drawing.Point(8, 384);
			this.ugtMEs.Name = "ugtMEs";
			this.ugtMEs.Size = new System.Drawing.Size(480, 216);
			this.ugtMEs.TabIndex = 964;
			this.ugtMEs.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugtMEs_ClickCellButton);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64});
			// 
			// frmReporteGastos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1232, 606);
			this.Controls.Add(this.ugtMEs);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.ugridPorcentaje);
			this.Controls.Add(this.btnAcciones);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.lblContador2);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnVer2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtDesde2);
			this.Controls.Add(this.dtHasta2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridGastos2);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.KeyPreview = true;
			this.Name = "frmReporteGastos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte De Gastos Anuales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteGastos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteGastos_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugtMEs)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		int iMes ;
		int iYear ;
		int iMesA ;
		int iYearA ;

		private void frmReporteGastos_Load(object sender, System.EventArgs e)
		{
			
			miAcceso = new Acceso(cdsSeteoF, "200524");
			if (miAcceso.BExportar) this.btnAcciones.Enabled =  true;
			Estado();

			#region Parte uno
			DateTime dtInicio = DateTime.Today.AddYears(-1);
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;	
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde2.Value = "01/" + DateTime.Now.Month.ToString() + "/" + dtInicio.Year.ToString();	
			this.dtHasta2.Value = DateTime.Today.AddYears(-1);

			#endregion Parte uno

			#region Parte dos 

			this.dtDesde2.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta2.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde2.CalendarInfo.MaxDate = DateTime.Today;			
//			this.dtHasta2.CalendarInfo.MaxDate = DateTime.Today;

			#endregion Parte dos 
		}

		private void ultraCalendarCombo2_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
     
			if ((int)this.optEstado.Value == 0)
			{
				if (this.dtHasta.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}
			}
			#endregion Valida Fechas

			if ( (int)this.optEstado.Value == 0)
			{
				string sSQL = string.Format("Exec PorcentajePresupuestoAnualGastos '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			}
			else
		{
			 iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Month));
			 iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtHasta.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			string sSQLPa = string.Format("Exec PorcentajePresupuestoAnualGastos '{0}', '{1}'", 
				Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), 	Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
		}
				this.lblContador.Text = this.uGridGastos.Rows.Count + " Registros Encontrados";

		}

		private void btnVer2_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde2.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde2.Focus();
				return;
			}
			if ((int)this.optEstado.Value == 0)
			{
				if (this.dtHasta2.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtDesde2.Value) > Convert.ToDateTime(this.dtHasta2.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}
			}
			#endregion Valida Fechas

			if ( (int)this.optEstado.Value == 0)
			{

				string sSQLG = string.Format("Exec PorcentajePresupuestoAnualGastos '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde2.Value).ToString("yyyyMMdd"), 	Convert.ToDateTime(this.dtHasta2.Value).ToString("yyyyMMdd"));
				this.uGridGastos2.DataSource = FuncionesProcedimientos.dtGeneral(sSQLG);
			}
			else
			{
				iMes = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Month));
				iYear = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Year));

				DateTime dtInicio = DateTime.Parse("01/" + iMes + "/" + iYear);
				DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
				string sSQLA = string.Format("Exec PorcentajePresupuestoAnualGastos '{0}', '{1}'", 
					Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), 	Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));
				this.uGridGastos2.DataSource = FuncionesProcedimientos.dtGeneral(sSQLA);
			}
				this.lblContador2.Text = this.uGridGastos2.Rows.Count + " Registros Encontrados";
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{

			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
			if ((int)this.optEstado.Value == 0)
			{
				if (this.dtHasta.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();
					return;
				}
			}
			#endregion Valida Fechas

			#region Valida Fechas
			if (this.dtDesde2.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde2.Focus();
				return;
			}
			if ((int)this.optEstado.Value == 0)
			{
				if (this.dtHasta2.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtDesde2.Value) > Convert.ToDateTime(this.dtHasta2.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}
			}
			#endregion Valida Fechas

			#region Periodo Repetido

			int icont = 0;
			if (Convert.ToDateTime(this.dtDesde.Value) == Convert.ToDateTime(this.dtDesde2.Value) )
				//				& Convert.ToDateTime(this.dtHasta.Value) = Convert.ToDateTime(this.dtHasta2.Value ))
			{
				icont = icont +1 ;
			}
			if ((int)this.optEstado.Value == 0)
			{

			if (Convert.ToDateTime(this.dtHasta.Value) == Convert.ToDateTime(this.dtHasta2.Value) )
				//				& Convert.ToDateTime(this.dtHasta.Value) = Convert.ToDateTime(this.dtHasta2.Value ))
			{
				icont = icont +1 ;
			}
			
				if (icont >= 2)
				{
					MessageBox.Show("Las Fechas del Periodo A y B Son las mismas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			
				}
			}
			else
			{
				if (icont == 1)
				{
					MessageBox.Show("Las Fechas del Periodo A y B Son las mismas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
     	#endregion Periodo Repetido

			if ( (int)this.optEstado.Value == 0)
			{
				string sSQLPorc = string.Format("Exec ConsultaPorcentajeAnualGastos '{0}', '{1}', '{2}', '{3}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtDesde2.Value).ToString("yyyyMMdd"), 	Convert.ToDateTime(this.dtHasta2.Value).ToString("yyyyMMdd"));
				this.ugridPorcentaje.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPorc);
			}
			else
			{
				int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Month));
				int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Year));

				int iMesMB = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Month));
				int iYearMB = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Year));

				DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
				DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			  
				DateTime dtInicioB = DateTime.Parse("01/" + iMesMB + "/" + iYearMB);
				DateTime dtFinalPeriodoB = DateTime.Parse(DateTime.DaysInMonth(dtInicioB.Date.Year, dtInicioB.Month).ToString() + "/" + dtInicioB.Month.ToString() + "/" + dtInicioB.Year.ToString());
			

				string sSQLPorcV = string.Format("Exec ConsultaPorcentajeMesGastos '{0}', '{1}', '{2}', '{3}'", 
					Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"),Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"),
					Convert.ToDateTime(dtInicioB).ToString("yyyyMMdd"), 	Convert.ToDateTime(dtFinalPeriodoB).ToString("yyyyMMdd"));
				this.ugtMEs.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPorcV);

			}
 		

		}

		private void uGridGastos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
			string sPeriodo = sColumna.Substring(3).ToUpper();
			
			if (e.Cell.Row.Cells[sPeriodo].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
      int iPeriodo = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).ToString("yyyy")));
			this.MostrarDetalle(e.Cell.Column.ToString(), (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, sPeriodo, e.Cell.Row.Cells["Grupo"].Value.ToString(), iPeriodo);
		}

		private void uGridGastos2_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
			string sPeriodo = sColumna.Substring(3).ToUpper();
			
			if (e.Cell.Row.Cells[sPeriodo].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
      int iPeriodo = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).ToString("yyyy")));
			this.MostrarDetalle(e.Cell.Column.ToString(), (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, sPeriodo, e.Cell.Row.Cells["Grupo"].Value.ToString(),iPeriodo);

		}


		private void MostrarDetalle(string sColumna, int idCuentaAgrupa, string sPeriodo, string sGasto, int IPeriodo)
		{			
			string sMes = "01/";
			if (sColumna == "btnFebrero") sMes = "02/";
			if (sColumna == "btnMarzo") sMes = "03/";
			if (sColumna == "btnAbril") sMes = "04/";
			if (sColumna == "btnMayo") sMes = "05/";
			if (sColumna == "btnJunio") sMes = "06/";
			if (sColumna == "btnJulio") sMes = "07/";
			if (sColumna == "btnAgosto") sMes = "08/";
			if (sColumna == "btnSeptiembre") sMes = "09/";
			if (sColumna == "btnOctubre") sMes = "10/";
			if (sColumna == "btnNoviembre") sMes = "11/";
			if (sColumna == "btnDiciembre") sMes = "12/";
			
			DateTime dtDesde = DateTime.Parse("01/" + sMes + IPeriodo.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());
			frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde, dtHasta, idCuentaAgrupa, sPeriodo, sGasto);
			prg.ShowDialog();


		}

		private void btnAcciones_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnAcciones, new Point(1, 25));
		}

		private void mnuAcreditacionTransferencias_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridGastos);
		}

		private void mnuAcreditacionCheques_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridGastos2);
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ugridPorcentaje);
		}

		private void frmReporteGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		
			
		

		private void ugridPorcentaje_ClickCellButton_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
			string sPeriodo = sColumna.Substring(3).ToUpper();

			#region Valida Fechas
			if (this.dtDesde2.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde2.Focus();
				return;
			}
			if ((int)this.optEstado.Value == 0)
			{
				if (this.dtHasta2.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtDesde2.Value) > Convert.ToDateTime(this.dtHasta2.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta2.Focus();
					return;
				}
			}
			#endregion Valida Fechas
			
			if (e.Cell.Row.Cells[sPeriodo].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((decimal)e.Cell.Row.Cells[sPeriodo].Value == 0.00m)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sMes = "01/";
			if (sColumna == "btnFebrero") sMes = "02/";
			if (sColumna == "btnMarzo") sMes = "03/";
			if (sColumna == "btnAbril") sMes = "04/";
			if (sColumna == "btnMayo") sMes = "05/";
			if (sColumna == "btnJunio") sMes = "06/";
			if (sColumna == "btnJulio") sMes = "07/";
			if (sColumna == "btnAgosto") sMes = "08/";
			if (sColumna == "btnSeptiembre") sMes = "09/";
			if (sColumna == "btnOctubre") sMes = "10/";
			if (sColumna == "btnNoviembre") sMes = "11/";
			if (sColumna == "btnDiciembre") sMes = "12/";

	

			int iPeriodo = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).ToString("yyyy")));
			int iPeriodo2 = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).ToString("yyyy")));
			
			DateTime dtDesde = DateTime.Parse("01/" + sMes + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			DateTime dtDesde2 = DateTime.Parse("01/" + sMes + iPeriodo2.ToString());
			DateTime dtHasta2 = DateTime.Parse(DateTime.DaysInMonth(dtDesde2.Date.Year, dtDesde2.Month).ToString() + "/" + dtDesde2.Month.ToString() + "/" + dtDesde2.Year.ToString());

			frmReporteComparacionGastos prg = new frmReporteComparacionGastos((DateTime)this.dtDesde.Value, (DateTime)this.dtHasta.Value, (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value , (DateTime)this.dtDesde2.Value, 
				(DateTime)this.dtHasta2.Value, e.Cell.Row.Cells["Grupo"].Value.ToString(), sPeriodo, true);
			prg.ShowDialog();

		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			Estado();
		}
		private void Estado ()
		{
			string sSQL = string.Format("Exec PorcentajePresupuestoAnualGastos '{0}', '{1}'", 
				"20000101","20000101");
			this.uGridGastos2.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			this.ugridPorcentaje.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			string sSQLb = string.Format("Exec [ConsultaPorcentajeMesGastos] '{0}', '{1}','{2}', '{3}'", 
				"20000101","20000101","20000101","20000101");
			this.ugtMEs.DataSource = FuncionesProcedimientos.dtGeneral(sSQLb);
			

			this.lblContador.Text ="";
			this.lblContador2.Text = "";

			

			if ( (int)this.optEstado.Value == 1)
			{
				this.dtDesde.Format = "MM/yyyy";
				this.dtHasta.Format = "MM/yyyy";
				this.dtDesde2.Format = "MM/yyyy";
				this.dtHasta2.Format = "MM/yyyy";

				this.dtHasta2.Enabled = false;
				this.dtHasta.Enabled = false;
				this.ugridPorcentaje.Visible = false;
				this.ugtMEs.Visible = true;

		
			}
			else
			{
				this.dtDesde.Format = "dd/MM/yyyy";
				this.dtHasta.Format = "dd/MM/yyyy";
				this.dtDesde2.Format = "dd/MM/yyyy";
				this.dtHasta2.Format = "dd/MM/yyyy";
				this.dtHasta2.Enabled = true;
				this.dtHasta.Enabled = true;
				this.ugridPorcentaje.Visible = true;
				this.ugtMEs.Visible = false;
			}
		
		}

		private void ugtMEs_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
			string sPeriodo = "Porcentaje";
			
			if (e.Cell.Row.Cells[sPeriodo].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((decimal)e.Cell.Row.Cells[sPeriodo].Value == 0.00m)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			string sMes = "ENERO";
	
			int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Month));
			int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde.Value).Year));

			int iMesMB = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Month));
			int iYearMB = Convert.ToInt32( ( Convert.ToDateTime(this.dtDesde2.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			  
			DateTime dtInicioB = DateTime.Parse("01/" + iMesMB + "/" + iYearMB);
			DateTime dtFinalPeriodoB = DateTime.Parse(DateTime.DaysInMonth(dtInicioB.Date.Year, dtInicioB.Month).ToString() + "/" + dtInicioB.Month.ToString() + "/" + dtInicioB.Year.ToString());

			frmReporteComparacionGastos prg = new frmReporteComparacionGastos(dtInicio, dtFinalPeriodo, (int)e.Cell.Row.Cells["idCuentaGrupo"].Value , dtInicioB, dtFinalPeriodoB, e.Cell.Row.Cells["Grupo"].Value.ToString(), sMes, false);
			prg.ShowDialog();
		}
	}
}
