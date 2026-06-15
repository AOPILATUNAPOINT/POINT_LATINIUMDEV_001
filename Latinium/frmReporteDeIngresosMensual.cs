using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteDeIngresosMensual.
	/// </summary>
	public class frmReporteDeIngresosMensual : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrValores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Button btnValores;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optDentalis;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnArea;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbControlable;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Button btnGrupos;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaGrupo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPromedio;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.Misc.UltraLabel lblPeriodo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFijo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArea;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGastoMes;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optVista;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeIngresosMensual()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 3, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeIngresosMensual));
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosControlableNo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosFijoVariable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosControlableNo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosFijoVariable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAreaGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAreaGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAreaGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosFijoVariable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosControlableNo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeFijo");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa", -1, "cmbCuentaGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAreaGastos", -1, "cmbArea");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosFijoVariable", -1, "cmbFijo", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosControlableNo", -1, "cmbControlable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeFijo");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			this.ugrValores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnValores = new System.Windows.Forms.Button();
			this.optDentalis = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnArea = new System.Windows.Forms.Button();
			this.cmbControlable = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.label53 = new System.Windows.Forms.Label();
			this.btnGrupos = new System.Windows.Forms.Button();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbCuentaGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.optPromedio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPeriodo = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFijo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbArea = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optGastoMes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optVista = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExcel = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ugrValores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optDentalis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbControlable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPromedio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFijo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastoMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			this.SuspendLayout();
			// 
			// ugrValores
			// 
			this.ugrValores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrValores.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrValores.DataSource = this.ultraDataSource6;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugrValores.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 230;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 15;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 4;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 6;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 7;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 9;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 11;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 12;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 13;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance14;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 14;
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
																										 ultraGridColumn16});
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
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
			ultraGridBand1.SummaryFooterCaption = "";
			this.ugrValores.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugrValores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugrValores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugrValores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.ugrValores.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugrValores.DisplayLayout.Override.CellAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 8F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugrValores.DisplayLayout.Override.HeaderAppearance = appearance30;
			this.ugrValores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrValores.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrValores.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.ugrValores.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugrValores.Location = new System.Drawing.Point(592, 7);
			this.ugrValores.Name = "ugrValores";
			this.ugrValores.Size = new System.Drawing.Size(680, 480);
			this.ugrValores.TabIndex = 778;
			this.ugrValores.Visible = false;
			this.ugrValores.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugrValores_KeyDown);
			this.ugrValores.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugrValores_AfterCellUpdate);
			this.ugrValores.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugrValores_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
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
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// btnValores
			// 
			this.btnValores.Image = ((System.Drawing.Image)(resources.GetObject("btnValores.Image")));
			this.btnValores.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnValores.Location = new System.Drawing.Point(376, 72);
			this.btnValores.Name = "btnValores";
			this.btnValores.Size = new System.Drawing.Size(88, 24);
			this.btnValores.TabIndex = 777;
			this.btnValores.Text = "Presupuesto";
			this.btnValores.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnValores.Visible = false;
			this.btnValores.Click += new System.EventHandler(this.btnValores_Click);
			// 
			// optDentalis
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.optDentalis.Appearance = appearance33;
			this.optDentalis.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optDentalis.CheckedIndex = 0;
			this.optDentalis.ItemAppearance = appearance34;
			this.optDentalis.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Detallado";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Agrupado";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Area";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "F/V";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "Contr/No Contr";
			this.optDentalis.Items.Add(valueListItem1);
			this.optDentalis.Items.Add(valueListItem2);
			this.optDentalis.Items.Add(valueListItem3);
			this.optDentalis.Items.Add(valueListItem4);
			this.optDentalis.Items.Add(valueListItem5);
			this.optDentalis.ItemSpacingVertical = 10;
			this.optDentalis.Location = new System.Drawing.Point(192, 40);
			this.optDentalis.Name = "optDentalis";
			this.optDentalis.Size = new System.Drawing.Size(376, 24);
			this.optDentalis.TabIndex = 776;
			this.optDentalis.Text = "Detallado";
			this.optDentalis.ValueChanged += new System.EventHandler(this.optDentalis_ValueChanged);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// btnArea
			// 
			this.btnArea.Enabled = false;
			this.btnArea.Image = ((System.Drawing.Image)(resources.GetObject("btnArea.Image")));
			this.btnArea.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnArea.Location = new System.Drawing.Point(448, 104);
			this.btnArea.Name = "btnArea";
			this.btnArea.Size = new System.Drawing.Size(75, 24);
			this.btnArea.TabIndex = 775;
			this.btnArea.Text = "Área";
			this.btnArea.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnArea.Click += new System.EventHandler(this.btnArea_Click);
			// 
			// cmbControlable
			// 
			this.cmbControlable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbControlable.DataSource = this.ultraDataSource5;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbControlable.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbControlable.DisplayMember = "Nombre";
			this.cmbControlable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbControlable.Location = new System.Drawing.Point(296, 224);
			this.cmbControlable.Name = "cmbControlable";
			this.cmbControlable.Size = new System.Drawing.Size(256, 112);
			this.cmbControlable.TabIndex = 774;
			this.cmbControlable.ValueMember = "idGastosControlableNo";
			this.cmbControlable.Visible = false;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(200, 10);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 771;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGrupos
			// 
			this.btnGrupos.Enabled = false;
			this.btnGrupos.Image = ((System.Drawing.Image)(resources.GetObject("btnGrupos.Image")));
			this.btnGrupos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrupos.Location = new System.Drawing.Point(360, 104);
			this.btnGrupos.Name = "btnGrupos";
			this.btnGrupos.Size = new System.Drawing.Size(75, 24);
			this.btnGrupos.TabIndex = 767;
			this.btnGrupos.Text = "Grupos";
			this.btnGrupos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrupos.Click += new System.EventHandler(this.btnGrupos_Click);
			// 
			// optConsulta
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance35;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 0;
			this.optConsulta.ItemAppearance = appearance36;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "Detallado";
			valueListItem7.DataValue = 2;
			valueListItem7.DisplayText = "Agrupado";
			this.optConsulta.Items.Add(valueListItem6);
			this.optConsulta.Items.Add(valueListItem7);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(8, 40);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(176, 24);
			this.optConsulta.TabIndex = 766;
			this.optConsulta.Text = "Detallado";
			this.optConsulta.ValueChanged += new System.EventHandler(this.optConsulta_ValueChanged);
			// 
			// cmbCuentaGrupo
			// 
			this.cmbCuentaGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaGrupo.DataSource = this.ultraDataSource2;
			ultraGridBand3.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 138;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 278;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCuentaGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaGrupo.DisplayMember = "Grupo";
			this.cmbCuentaGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaGrupo.Location = new System.Drawing.Point(40, 264);
			this.cmbCuentaGrupo.Name = "cmbCuentaGrupo";
			this.cmbCuentaGrupo.Size = new System.Drawing.Size(280, 112);
			this.cmbCuentaGrupo.TabIndex = 765;
			this.cmbCuentaGrupo.ValueMember = "idCuentaAgrupa";
			this.cmbCuentaGrupo.Visible = false;
			// 
			// optPromedio
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPromedio.Appearance = appearance37;
			this.optPromedio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPromedio.CheckedIndex = 0;
			this.optPromedio.ItemAppearance = appearance38;
			this.optPromedio.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 1;
			valueListItem8.DisplayText = "Anual";
			valueListItem9.DataValue = 2;
			valueListItem9.DisplayText = "Ultimo semestre";
			this.optPromedio.Items.Add(valueListItem8);
			this.optPromedio.Items.Add(valueListItem9);
			this.optPromedio.ItemSpacingVertical = 10;
			this.optPromedio.Location = new System.Drawing.Point(8, 104);
			this.optPromedio.Name = "optPromedio";
			this.optPromedio.Size = new System.Drawing.Size(176, 24);
			this.optPromedio.TabIndex = 762;
			this.optPromedio.Text = "Anual";
			this.optPromedio.ValueChanged += new System.EventHandler(this.optPromedio_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(200, 104);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 758;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtAnio
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance39;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(472, 8);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 780;
			this.txtAnio.Visible = false;
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.BackColor = System.Drawing.Color.Transparent;
			this.lblPeriodo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblPeriodo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblPeriodo.Location = new System.Drawing.Point(432, 11);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(23, 15);
			this.lblPeriodo.TabIndex = 779;
			this.lblPeriodo.Text = "Año";
			this.lblPeriodo.Visible = false;
			// 
			// cmbFijo
			// 
			this.cmbFijo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFijo.DataSource = this.ultraDataSource4;
			ultraGridBand4.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbFijo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFijo.DisplayMember = "Nombre";
			this.cmbFijo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFijo.Location = new System.Drawing.Point(296, 296);
			this.cmbFijo.Name = "cmbFijo";
			this.cmbFijo.Size = new System.Drawing.Size(256, 112);
			this.cmbFijo.TabIndex = 773;
			this.cmbFijo.ValueMember = "idGastosFijoVariable";
			this.cmbFijo.Visible = false;
			// 
			// cmbArea
			// 
			this.cmbArea.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArea.DataSource = this.ultraDataSource3;
			ultraGridBand5.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 250;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbArea.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbArea.DisplayMember = "Nombre";
			this.cmbArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArea.Location = new System.Drawing.Point(40, 224);
			this.cmbArea.Name = "cmbArea";
			this.cmbArea.Size = new System.Drawing.Size(280, 112);
			this.cmbArea.TabIndex = 772;
			this.cmbArea.ValueMember = "idAreaGastos";
			this.cmbArea.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 770;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance40;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 768;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance41;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 769;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// optGastoMes
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGastoMes.Appearance = appearance42;
			this.optGastoMes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGastoMes.CheckedIndex = 0;
			this.optGastoMes.ItemAppearance = appearance43;
			this.optGastoMes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem10.DataValue = 1;
			valueListItem10.DisplayText = "Por Gasto";
			valueListItem11.DataValue = 2;
			valueListItem11.DisplayText = "Por Mes";
			this.optGastoMes.Items.Add(valueListItem10);
			this.optGastoMes.Items.Add(valueListItem11);
			this.optGastoMes.ItemSpacingVertical = 10;
			this.optGastoMes.Location = new System.Drawing.Point(192, 72);
			this.optGastoMes.Name = "optGastoMes";
			this.optGastoMes.Size = new System.Drawing.Size(176, 24);
			this.optGastoMes.TabIndex = 764;
			this.optGastoMes.Text = "Por Gasto";
			this.optGastoMes.ValueChanged += new System.EventHandler(this.optGastoMes_ValueChanged);
			// 
			// optVista
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.optVista.Appearance = appearance44;
			this.optVista.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optVista.CheckedIndex = 0;
			this.optVista.ItemAppearance = appearance45;
			this.optVista.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem12.DataValue = 1;
			valueListItem12.DisplayText = "Valores";
			valueListItem13.DataValue = 2;
			valueListItem13.DisplayText = "Porcentajes";
			this.optVista.Items.Add(valueListItem12);
			this.optVista.Items.Add(valueListItem13);
			this.optVista.ItemSpacingVertical = 10;
			this.optVista.Location = new System.Drawing.Point(8, 72);
			this.optVista.Name = "optVista";
			this.optVista.Size = new System.Drawing.Size(176, 24);
			this.optVista.TabIndex = 763;
			this.optVista.Text = "Valores";
			this.optVista.ValueChanged += new System.EventHandler(this.optVista_ValueChanged);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance46.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance46;
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance49.FontData.BoldAsString = "True";
			appearance49.FontData.Name = "Arial";
			appearance49.FontData.SizeInPoints = 8F;
			appearance49.ForeColor = System.Drawing.Color.White;
			appearance49.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance49;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance51;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(592, 8);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(680, 480);
			this.uGridInformacion.TabIndex = 761;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(bool);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35});
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
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance52.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance52;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 52;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 55;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Cuenta";
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Width = 113;
			ultraGridColumn28.Header.Caption = "G";
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Width = 18;
			ultraGridColumn29.Header.Caption = "Grupo";
			ultraGridColumn29.Header.VisiblePosition = 4;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn29.Width = 102;
			ultraGridColumn30.Header.Caption = "F-V-E";
			ultraGridColumn30.Header.VisiblePosition = 5;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 40;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance53;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "% Variable";
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn31.Width = 71;
			ultraGridColumn32.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn32.Header.Caption = "Área";
			ultraGridColumn32.Header.VisiblePosition = 6;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn32.Width = 120;
			ultraGridColumn33.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn33.Header.Caption = "F/V";
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn33.Width = 120;
			ultraGridColumn34.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn34.Header.Caption = "Contr/No Contr";
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn34.Width = 120;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance54;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "% Fijo";
			ultraGridColumn35.Header.VisiblePosition = 10;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			ultraGridBand7.Header.Caption = "Gastos";
			ultraGridBand7.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance55.ForeColor = System.Drawing.Color.Black;
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance55;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance56.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 8F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance57;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 136);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(576, 352);
			this.uGridGastos.TabIndex = 760;
			this.uGridGastos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_CellChange);
			this.uGridGastos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridGastos_KeyPress);
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(280, 104);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 759;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.IndianRed;
			this.label1.Location = new System.Drawing.Point(480, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 781;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Visible = false;
			// 
			// frmReporteDeIngresosMensual
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1280, 494);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.optPromedio);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.lblPeriodo);
			this.Controls.Add(this.cmbFijo);
			this.Controls.Add(this.cmbArea);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.optGastoMes);
			this.Controls.Add(this.optVista);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.ugrValores);
			this.Controls.Add(this.btnValores);
			this.Controls.Add(this.optDentalis);
			this.Controls.Add(this.btnArea);
			this.Controls.Add(this.cmbControlable);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.btnGrupos);
			this.Controls.Add(this.optConsulta);
			this.Controls.Add(this.cmbCuentaGrupo);
			this.Name = "frmReporteDeIngresosMensual";
			this.Text = "Reporte Ingresos Mensual";
			this.Load += new System.EventHandler(this.frmReporteDeIngresosMensual_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugrValores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optDentalis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbControlable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPromedio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFijo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGastoMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optVista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

			private Acceso miAcceso;
		private void frmReporteDeIngresosMensual_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso (cdsSeteoF,"47003");
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.Nuevo)
			{
				this.btnGrupos.Enabled = true;
				this.btnArea.Enabled = true;
			}
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");


			string sSQL = string.Format("Select idCuenta, Codigo, Descripcion, Gasto, idGastosCuentaGrupo as idCuentaAgrupa, TipoGasto, Porcentaje, ISNULL (idAreaGastos, 0) AS idAreaGastos , ISNULL (idGastosFijoVariable, 0) AS idGastosFijoVariable , ISNULL (idGastosControlableNo, 0) AS idGastosControlableNo, PorcentajeFijo From Cuenta Where Tipo = 4 And Grupo = 0 Order By Codigo");
			FuncionesProcedimientos.dsGeneralModulo (sSQL, this.uGridGastos);

				this.cmbCuentaGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idGastosCuentaGrupo as idCuentaAgrupa, Grupo From GastosCuentaGrupo Order By Grupo"); 
			this.cmbArea.DataSource = FuncionesProcedimientos.dtGeneral("select idAreaGastos, Nombre from AreaGastos order by Nombre");
			this.cmbControlable.DataSource = FuncionesProcedimientos.dtGeneral("select idGastosControlableNo, Nombre from GastosControlableNo order by Nombre");
			this.cmbFijo.DataSource = FuncionesProcedimientos.dtGeneral("select idGastosFijoVariable, Nombre from GastosFijoVariable order by Nombre");

			//this.cmbFijo.DataSource = FuncionesProcedimientos.dtGeneral("");
			this.optConsulta.Visible = false;
			this.optDentalis.Location = new System.Drawing.Point(8, 40);
			if (miAcceso.Nuevo) this.btnValores.Visible = true;
			this.txtAnio.MinValue = 2021;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
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
		private void optConsulta_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void optDentalis_ValueChanged(object sender, System.EventArgs e)
		{
				this.btnVer_Click (sender, e);
		}

		private void optVista_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void optGastoMes_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void optPromedio_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		private void btnValores_Click(object sender, System.EventArgs e)
		{
			this.ugrValores.Visible = true;
			this.uGridInformacion.Visible = false;
			this.lblPeriodo.Visible = true;
			this.txtAnio.Visible = true;
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2020, DateTime.Today.Year + 1)) return;
			string sGastos = "";						

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{	
				if ((bool)dr.Cells["Gasto"].Value)
				{
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString(); 
					else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();
				}				
			}

			string sSQL = string.Format("Exec [ConsultaPresupuestoIngreso] {0}", (int) this.txtAnio.Value );
			this.ugrValores.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idCuentaAgrupa")
			{									
				int idCuentaAgrupa = 0;
				if (e.Cell.Row.Cells["idCuentaAgrupa"].Value != System.DBNull.Value) idCuentaAgrupa = (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idGastosCuentaGrupo = {0} Where idCuenta = {1}", idCuentaAgrupa, (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	
		
//			if (e.Cell.Column.ToString() == "TipoGasto")
//				if (e.Cell.Row.Cells["TipoGasto"].Value.ToString() != "V")
//					e.Cell.Row.Cells["Porcentaje"].Value = 0.00m;

			if (e.Cell.Column.ToString() == "TipoGasto")
			{					
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set TipoGasto = '{0}' Where idCuenta = {1}", e.Cell.Row.Cells["TipoGasto"].Value.ToString(), (bool)e.Cell.Row.Cells["Fijo"].Value));
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set Porcentaje = {0} Where idCuenta = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Porcentaje"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	
			if (e.Cell.Column.ToString() == "PorcentajeFijo")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set PorcentajeFijo = {0} Where idCuenta = {1}", Convert.ToDecimal(e.Cell.Row.Cells["PorcentajeFijo"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	

			if (e.Cell.Column.ToString() == "idAreaGastos")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idAreaGastos = {0} Where idCuenta = {1}", Convert.ToInt32(e.Cell.Row.Cells["idAreaGastos"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	

			if (e.Cell.Column.ToString() == "idGastosControlableNo")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idGastosControlableNo = {0} Where idCuenta = {1}", Convert.ToInt32(e.Cell.Row.Cells["idGastosControlableNo"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	

			if (e.Cell.Column.ToString() == "idGastosFijoVariable")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idGastosFijoVariable = {0} Where idCuenta = {1}", Convert.ToInt32(e.Cell.Row.Cells["idGastosFijoVariable"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
			}	

			if (this.btnValores.Enabled)
			{
				if (e.Cell.Column.ToString() == "Gasto")
				{
					if ((bool)e.Cell.Row.Cells["Gasto"].Value)
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [GeneraPeriodoIngresoAnual] {0}, {1}", (int) this.txtAnio.Value ,(int)e.Cell.Row.Cells["idCuenta"].Value));  
					}
				}
			}
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
			#endregion Valida Fechas

			string sGastos = "";						

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{	
				if ((bool)dr.Cells["Gasto"].Value)
				{
					if (sGastos.Length == 0) sGastos = dr.Cells["idCuenta"].Value.ToString(); 
					else if (sGastos.Length > 0) sGastos = sGastos + ", " + dr.Cells["idCuenta"].Value.ToString();
				}				
			}

			this.uGridInformacion.DataSource = null;
			
			if (MenuLatinium.stDirFacturacion == "POINT" || MenuLatinium.stDirFacturacion == "CREDISOLUTIONS")
			{
				string sProcedimiento = "ReporteDeIngresoAnual";
				if ((int)this.optConsulta.Value > 1) sProcedimiento = "ReporteDeIngresosAnualAgrupado";

				string sSQL = string.Format("Exec {0} '{1}', '{2}', '{3}', {4}", sProcedimiento, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sGastos, (int)this.optConsulta.Value);
				this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				this.DiseñoGrid();
			}
			else
			{
				string sProcedimiento = "ReporteDeGastosAnual";
				if ((int)this.optDentalis.Value > 1) sProcedimiento = "ReporteDeGastosAnualAgrupado";
//pendiente
//				string sSQL = string.Format("Exec {0} '{1}', '{2}', '{3}', {4}", sProcedimiento, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sGastos, (int)this.optDentalis.Value);
//				this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
//				this.DiseñoGridDentalis();
			}
		}

		private void uGridGastos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value)
			{
				
				if (e.Cell.Column.ToString() == "GASTO")
				{
					this.uGridGastos.UpdateData();
					
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaGastoCuentas {0}, {1}", (int)e.Cell.Row.Cells["idCuenta"].Value, (bool)e.Cell.Row.Cells["Gasto"].Value));
				}				
				if (this.btnValores.Enabled)
				{
					if (e.Cell.Column.ToString() == "Gasto")
					{
						if ((bool)e.Cell.Row.Cells["Gasto"].Value)
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [GeneraPeriodoIngresoAnual] {0}, {1}", (int) this.txtAnio.Value ,(int)e.Cell.Row.Cells["idCuenta"].Value));  
						}
					}
				}
			}
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
				FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
		}
		
		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 300;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Activation.ActivateOnly;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Semestre", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Semestre"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Semestre"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].Width = 90;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Width = 90;
				if ((int)this.optVista.Value == 2 && (int)this.optGastoMes.Value == 2) this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = false;
				else this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellActivation = Activation.NoEdit;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"].CellAppearance.BackColor = Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("PromedioMes", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["PromedioMes"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PromedioMes"].DisplayFormat = "{0: #,##0.00}";
				
				if (this.uGridInformacion.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 6; j < iColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}

					string sNCSuma = "";
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						int iPromedioAnual = 0;
						int iPromedioSemestral = 0;

						for (int i = 6; i < iColumnas; i++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[i].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) iPromedioAnual++;
							if (i > 11) if (dr.Cells[sNCSuma].Value != System.DBNull.Value) iPromedioSemestral++;
						}

						for (int j = 6; j < iColumnas; j++)
						{
							sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
							if (dr.Cells[sNCSuma].Value != System.DBNull.Value) dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							if (j > 11)
								if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									dr.Cells["Semestre"].Value = Convert.ToDecimal(dr.Cells["Semestre"].Value) + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
							              	
							if ((int)this.optPromedio.Value == 1)
							{
								if ( iPromedioAnual > 0)
								{
									dr.Cells["Promedio"].Value = Decimal.Round(Convert.ToDecimal(dr.Cells["Total"].Value) / iPromedioAnual, 2);
								}
								else
								{
									dr.Cells["Promedio"].Value = 0;
								}
							}
							if ((int)this.optPromedio.Value == 2)
							{
								if ( iPromedioSemestral > 0)
								{
									dr.Cells["Promedio"].Value = Decimal.Round(Convert.ToDecimal(dr.Cells["Semestre"].Value) / iPromedioSemestral, 2);
								}
								else
								{
									dr.Cells["Promedio"].Value = 0;
								}
							}
						}
					}

					sNCSuma = "";
					if ((int)optVista.Value == 2)
					{
						int iPromedioMes = 0;
						decimal dSumaTotalMes = 0.00m;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
						{
							iPromedioMes = 0;
							dSumaTotalMes = 0.00m;

							for (int j = 6; j < iColumnas; j++)
							{
								sNCSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;

								#region Porcentaje Por Gasto
								if ((int)optGastoMes.Value == 1)
								{
									if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									{
										if (Convert.ToDecimal(dr.Cells["Promedio"].Value) > 0.00m)
											if (Convert.ToDecimal(dr.Cells[sNCSuma].Value) > 0.00m)
												dr.Cells[sNCSuma].Value = Decimal.Round(Convert.ToDecimal(dr.Cells[sNCSuma].Value) * 100/Convert.ToDecimal(dr.Cells["Promedio"].Value), 2);
									}
								}
								#endregion Porcentaje Por Gasto

								#region Porcentaje Por Mes
								if ((int)optGastoMes.Value == 2)
								{
									decimal dTotalMes = 0.00m;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drTotal in this.uGridInformacion.Rows.All)
									{
										if (drTotal.Cells[sNCSuma].Value != System.DBNull.Value) 
											dTotalMes = dTotalMes + Convert.ToDecimal(drTotal.Cells[sNCSuma].Value);
									}
									
									if (dr.Cells[sNCSuma].Value != System.DBNull.Value) 
									{
										if (dTotalMes > 0.00m)
											if (Convert.ToDecimal(dr.Cells[sNCSuma].Value) > 0.00m)
												dr.Cells[sNCSuma].Value = Decimal.Round(Convert.ToDecimal(dr.Cells[sNCSuma].Value) * 100/dTotalMes, 2);

										iPromedioMes++;
										dSumaTotalMes = dSumaTotalMes + Convert.ToDecimal(dr.Cells[sNCSuma].Value);
									}
								}

								if (dSumaTotalMes > 0.00m)
									if (iPromedioMes > 0)
										dr.Cells["PromedioMes"].Value = Decimal.Round(dSumaTotalMes / iPromedioMes, 2);
								#endregion Porcentaje Por Mes
							}
						}
					}
				}				
			}

			this.uGridInformacion.DisplayLayout.Bands[0].Header.Caption = this.uGridInformacion.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
		}

		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGastos);
		}

		private void uGridGastos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			if (this.uGridGastos.ActiveRow.Cells["TipoGasto"].Value.ToString() != "V" && 
//				this.uGridGastos.ActiveCell.Column.Key.ToString().ToUpper() == "PORCENTAJE") e.Handled = true;
		}

		private void ugrValores_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores ( Convert.ToInt32 (e.Cell.Row.Cells["idCuenta"].Value), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}
		private void ActualizaValores (int idCuenta, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec [Com_ActualizaValoresPresupuestoIngreso] {0}, {1}, '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), idCuenta, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ugrValores_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

		private void ugrValores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugrValores);
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (MenuLatinium.stDirFacturacion == "POINT" || MenuLatinium.stDirFacturacion == "CREDISOLUTIONS")
			{
				if ( (int) this.optConsulta.Value == 1  || (int) this.optPromedio.Value == 2) return;
			}
			else
			{
				if ( (int) this.optDentalis.Value == 1 || (int) this.optPromedio.Value == 2) return;
			}
			string sColumna = e.Cell.Column.Key;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (e.Cell.Column.ToString() == sColumna)
			{
				if (MenuLatinium.stDirFacturacion == "POINT" || MenuLatinium.stDirFacturacion == "CREDISOLUTIONS")
				{
					using (frmReporteDeGastosAnualDetalle CCJ = new frmReporteDeGastosAnualDetalle ( (string)e.Cell.Row.Cells["Codigo"].Value, 
									 (string)e.Cell.Row.Cells["Descripcion"].Value , sColumna, (int) this.optConsulta.Value ))
					{
		
						if (DialogResult.OK == CCJ.ShowDialog())
						{
						
						}
					}
				}
				else
				{
					using (frmReporteDeGastosAnualDetalle CCJ = new frmReporteDeGastosAnualDetalle ( (string)e.Cell.Row.Cells["Codigo"].Value, 
									 (string)e.Cell.Row.Cells["Descripcion"].Value , sColumna, (int) this.optDentalis.Value ))
					{
		
						if (DialogResult.OK == CCJ.ShowDialog())
						{
						
						}
					}
				}
			}
		}

		private void btnGrupos_Click(object sender, System.EventArgs e)
		{
			using (frmCuentaGrupo CG = new frmCuentaGrupo(1))
			{				
				if (DialogResult.OK == CG.ShowDialog())
				{
					this.cmbCuentaGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idGastosCuentaGrupo idCuentaAgrupa, Grupo From GastosCuentaGrupo Order By Grupo");
				}
			}
		}

		private void btnArea_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}
	}
}
