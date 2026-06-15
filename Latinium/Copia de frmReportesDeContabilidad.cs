using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportesDeContabilidad.
	/// </summary>
	public class frmReportesDeContabilidad : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraTreeList.TreeList tlsCuentas;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCodigo;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colDescripcion;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colSaldoInicial;
		private DevExpress.XtraTreeList.Columns.TreeListColumn ColDebe;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colHaber;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colSaldoFinal;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFinal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInicial;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.SimpleButton btnImprimir;
		private DevExpress.XtraEditors.SimpleButton btnAsiento;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalHaber;
		private System.Windows.Forms.Label lblIce;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalDebe;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtMayorHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtMayorDesde;
		private DevExpress.XtraEditors.SimpleButton btConsultar;
		private DevExpress.XtraEditors.SimpleButton btnGeneraBalance;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private DevExpress.XtraEditors.SimpleButton btnImprimirBG;
		private DevExpress.XtraEditors.SimpleButton btnImprimirBR;
		private DevExpress.XtraEditors.SimpleButton btnImprimirBC;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBCDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBCHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBGDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBGHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBRDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtBRHasta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private DevExpress.XtraEditors.SimpleButton btnGeneraBalanceGeneral;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit1;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit2;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit3;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit4;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit1;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit2;
		private DevExpress.XtraTreeList.TreeList tlsBalanceGeneral;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn1;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit3;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn2;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit4;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn3;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit5;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn4;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit6;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn5;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit7;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn6;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit8;
		private DevExpress.XtraTreeList.TreeList tlsBalanceDeResultados;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn7;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit5;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn8;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit6;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn9;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit9;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn10;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit10;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn11;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit11;
		private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn12;
		private DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit repositoryItemSpinEdit12;
		private DevExpress.XtraEditors.SimpleButton btnGenerarBR;
		private DevExpress.XtraTreeList.Columns.TreeListColumn Grupo;
		private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBCFinal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBCInicial;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaFinal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaInicial;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private DevExpress.XtraEditors.SimpleButton btnGenerar;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.SimpleButton btnVacio;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpMayorGeneralCondensado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMayorCondensado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportesDeContabilidad()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoInicial");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoIni");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "SaldoInicial", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoInicial", 9, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 11, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 12, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoIni");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportesDeContabilidad));
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			DevExpress.XtraTreeList.StyleFormatConditions.StyleFormatCondition styleFormatCondition1 = new DevExpress.XtraTreeList.StyleFormatConditions.StyleFormatCondition();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton7 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton8 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoInicial");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "SaldoInicial", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoInicial", 3, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 4, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 5, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta5");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton9 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton10 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label6 = new System.Windows.Forms.Label();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtTotalHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIce = new System.Windows.Forms.Label();
			this.txtTotalDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtFinal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtInicial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtMayorHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtMayorDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btConsultar = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.tlsCuentas = new DevExpress.XtraTreeList.TreeList();
			this.colCodigo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.colDescripcion = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.colSaldoInicial = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.ColDebe = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.colHaber = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit3 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.colSaldoFinal = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit4 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.Grupo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.txtBCFinal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBCInicial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.dtBCDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtBCHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGeneraBalance = new DevExpress.XtraEditors.SimpleButton();
			this.label2 = new System.Windows.Forms.Label();
			this.btnImprimirBC = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.tlsBalanceGeneral = new DevExpress.XtraTreeList.TreeList();
			this.treeListColumn1 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit3 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.treeListColumn2 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit4 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.treeListColumn3 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit5 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn4 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit6 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn5 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit7 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn6 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit8 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.btnGeneraBalanceGeneral = new DevExpress.XtraEditors.SimpleButton();
			this.label8 = new System.Windows.Forms.Label();
			this.dtBGDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtBGHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.btnImprimirBG = new DevExpress.XtraEditors.SimpleButton();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.tlsBalanceDeResultados = new DevExpress.XtraTreeList.TreeList();
			this.treeListColumn7 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit5 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.treeListColumn8 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemTextEdit6 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.treeListColumn9 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit9 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn10 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit10 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn11 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit11 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.treeListColumn12 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemSpinEdit12 = new DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit();
			this.btnGenerarBR = new DevExpress.XtraEditors.SimpleButton();
			this.label10 = new System.Windows.Forms.Label();
			this.dtBRDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtBRHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.btnImprimirBR = new DevExpress.XtraEditors.SimpleButton();
			this.utpMayorGeneralCondensado = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridMayorCondensado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.btnImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btnAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.label14 = new System.Windows.Forms.Label();
			this.txtCuentas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.txtCuentaFinal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCuentaInicial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.btnGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnVacio = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalDebe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMayorHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMayorDesde)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBCFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBCInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBCDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBCHasta)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.tlsBalanceGeneral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBGDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBGHasta)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.tlsBalanceDeResultados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBRDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBRHasta)).BeginInit();
			this.utpMayorGeneralCondensado.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridMayorCondensado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.label6);
			this.ultraTabPageControl9.Controls.Add(this.txtSaldoTotal);
			this.ultraTabPageControl9.Controls.Add(this.label7);
			this.ultraTabPageControl9.Controls.Add(this.txtTotalHaber);
			this.ultraTabPageControl9.Controls.Add(this.lblIce);
			this.ultraTabPageControl9.Controls.Add(this.txtTotalDebe);
			this.ultraTabPageControl9.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl9.Controls.Add(this.label5);
			this.ultraTabPageControl9.Controls.Add(this.txtCuenta);
			this.ultraTabPageControl9.Controls.Add(this.lblComprobante);
			this.ultraTabPageControl9.Controls.Add(this.lblNumero);
			this.ultraTabPageControl9.Controls.Add(this.txtFinal);
			this.ultraTabPageControl9.Controls.Add(this.txtInicial);
			this.ultraTabPageControl9.Controls.Add(this.dtMayorHasta);
			this.ultraTabPageControl9.Controls.Add(this.dtMayorDesde);
			this.ultraTabPageControl9.Controls.Add(this.label3);
			this.ultraTabPageControl9.Controls.Add(this.label4);
			this.ultraTabPageControl9.Controls.Add(this.btConsultar);
			this.ultraTabPageControl9.Controls.Add(this.simpleButton1);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(2, 21);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1188, 417);
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(990, 390);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(64, 16);
			this.label6.TabIndex = 344;
			this.label6.Text = "Saldo Total ";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoTotal
			// 
			this.txtSaldoTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.BackColorDisabled = System.Drawing.Color.White;
			appearance1.BackColorDisabled2 = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance1;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(1070, 388);
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(112, 21);
			this.txtSaldoTotal.TabIndex = 343;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(774, 390);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(63, 16);
			this.label7.TabIndex = 342;
			this.label7.Text = "Total Haber";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalHaber
			// 
			this.txtTotalHaber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.BackColorDisabled = System.Drawing.Color.White;
			appearance2.BackColorDisabled2 = System.Drawing.Color.White;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalHaber.Appearance = appearance2;
			this.txtTotalHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalHaber.Enabled = false;
			this.txtTotalHaber.FormatString = "#,##0.00";
			this.txtTotalHaber.Location = new System.Drawing.Point(854, 388);
			this.txtTotalHaber.Name = "txtTotalHaber";
			this.txtTotalHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalHaber.PromptChar = ' ';
			this.txtTotalHaber.Size = new System.Drawing.Size(112, 21);
			this.txtTotalHaber.TabIndex = 341;
			// 
			// lblIce
			// 
			this.lblIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIce.AutoSize = true;
			this.lblIce.BackColor = System.Drawing.Color.Transparent;
			this.lblIce.Location = new System.Drawing.Point(558, 390);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(60, 16);
			this.lblIce.TabIndex = 340;
			this.lblIce.Text = "Total Debe";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalDebe
			// 
			this.txtTotalDebe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.BackColorDisabled = System.Drawing.Color.White;
			appearance3.BackColorDisabled2 = System.Drawing.Color.White;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalDebe.Appearance = appearance3;
			this.txtTotalDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalDebe.Enabled = false;
			this.txtTotalDebe.FormatString = "#,##0.00";
			this.txtTotalDebe.Location = new System.Drawing.Point(638, 388);
			this.txtTotalDebe.Name = "txtTotalDebe";
			this.txtTotalDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalDebe.PromptChar = ' ';
			this.txtTotalDebe.Size = new System.Drawing.Size(112, 21);
			this.txtTotalDebe.TabIndex = 339;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 79;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 72;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 97;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 187;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 61;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Numero";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "Referencia";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 70;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Nombre";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 181;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Saldo Inicial";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Saldo Inicial";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 112;
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
			ultraGridBand1.Header.Caption = "";
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnHeaderAppearance = appearance11;
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.Select;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.Hidden = true;
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance13;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance14;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1174, 368);
			this.ultraGrid1.TabIndex = 8;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn14.DefaultValue = 0;
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
																																 ultraDataColumn15});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(392, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(41, 16);
			this.label5.TabIndex = 336;
			this.label5.Text = "Cuenta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuenta
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuenta.Appearance = appearance22;
			this.txtCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuenta.Location = new System.Drawing.Point(448, 38);
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(216, 21);
			this.txtCuenta.TabIndex = 330;
			this.txtCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuenta_KeyDown);
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(200, 40);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(29, 16);
			this.lblComprobante.TabIndex = 335;
			this.lblComprobante.Text = "Final";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(8, 40);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(33, 16);
			this.lblNumero.TabIndex = 334;
			this.lblNumero.Text = "Inicial";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFinal
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFinal.Appearance = appearance23;
			this.txtFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFinal.Enabled = false;
			this.txtFinal.Location = new System.Drawing.Point(256, 38);
			this.txtFinal.Name = "txtFinal";
			this.txtFinal.Size = new System.Drawing.Size(120, 21);
			this.txtFinal.TabIndex = 329;
			this.txtFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFinal_KeyDown);
			// 
			// txtInicial
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInicial.Appearance = appearance24;
			this.txtInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInicial.Location = new System.Drawing.Point(64, 38);
			this.txtInicial.Name = "txtInicial";
			this.txtInicial.Size = new System.Drawing.Size(120, 21);
			this.txtInicial.TabIndex = 328;
			this.txtInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInicial_KeyDown);
			this.txtInicial.TextChanged += new System.EventHandler(this.txtInicial_TextChanged);
			this.txtInicial.ValueChanged += new System.EventHandler(this.txtInicial_ValueChanged);
			// 
			// dtMayorHasta
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMayorHasta.Appearance = appearance25;
			this.dtMayorHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtMayorHasta.DateButtons.Add(dateButton1);
			this.dtMayorHasta.Format = "dd/MM/yyyy";
			this.dtMayorHasta.Location = new System.Drawing.Point(256, 8);
			this.dtMayorHasta.Name = "dtMayorHasta";
			this.dtMayorHasta.NonAutoSizeHeight = 23;
			this.dtMayorHasta.Size = new System.Drawing.Size(112, 21);
			this.dtMayorHasta.SpinButtonsVisible = true;
			this.dtMayorHasta.TabIndex = 327;
			this.dtMayorHasta.Value = ((object)(resources.GetObject("dtMayorHasta.Value")));
			this.dtMayorHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtMayorHasta_KeyDown);
			// 
			// dtMayorDesde
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMayorDesde.Appearance = appearance26;
			this.dtMayorDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtMayorDesde.DateButtons.Add(dateButton2);
			this.dtMayorDesde.Format = "dd/MM/yyyy";
			this.dtMayorDesde.Location = new System.Drawing.Point(64, 8);
			this.dtMayorDesde.Name = "dtMayorDesde";
			this.dtMayorDesde.NonAutoSizeHeight = 23;
			this.dtMayorDesde.Size = new System.Drawing.Size(112, 21);
			this.dtMayorDesde.SpinButtonsVisible = true;
			this.dtMayorDesde.TabIndex = 326;
			this.dtMayorDesde.Value = ((object)(resources.GetObject("dtMayorDesde.Value")));
			this.dtMayorDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtMayorDesde_KeyDown);
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
			// btConsultar
			// 
			this.btConsultar.Location = new System.Drawing.Point(680, 38);
			this.btConsultar.Name = "btConsultar";
			this.btConsultar.Size = new System.Drawing.Size(80, 21);
			this.btConsultar.TabIndex = 331;
			this.btConsultar.Text = "Generar";
			this.btConsultar.Click += new System.EventHandler(this.btConsultar_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.Location = new System.Drawing.Point(680, 32);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(80, 21);
			this.simpleButton1.TabIndex = 644;
			this.simpleButton1.Text = "Asiento";
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.tlsCuentas);
			this.ultraTabPageControl10.Controls.Add(this.label12);
			this.ultraTabPageControl10.Controls.Add(this.label13);
			this.ultraTabPageControl10.Controls.Add(this.txtBCFinal);
			this.ultraTabPageControl10.Controls.Add(this.txtBCInicial);
			this.ultraTabPageControl10.Controls.Add(this.label1);
			this.ultraTabPageControl10.Controls.Add(this.dtBCDesde);
			this.ultraTabPageControl10.Controls.Add(this.dtBCHasta);
			this.ultraTabPageControl10.Controls.Add(this.btnGeneraBalance);
			this.ultraTabPageControl10.Controls.Add(this.label2);
			this.ultraTabPageControl10.Controls.Add(this.btnImprimirBC);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(1189, 418);
			this.ultraTabPageControl10.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl10_Paint);
			// 
			// tlsCuentas
			// 
			this.tlsCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsCuentas.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.tlsCuentas.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																															this.colCodigo,
																																															this.colDescripcion,
																																															this.colSaldoInicial,
																																															this.ColDebe,
																																															this.colHaber,
																																															this.colSaldoFinal,
																																															this.Grupo});
			styleFormatCondition1.Appearance.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold);
			styleFormatCondition1.Appearance.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
			styleFormatCondition1.Appearance.Options.UseFont = true;
			styleFormatCondition1.ApplyToRow = true;
			styleFormatCondition1.Condition = DevExpress.XtraGrid.FormatConditionEnum.Equal;
			styleFormatCondition1.Value1 = true;
			this.tlsCuentas.FormatConditions.AddRange(new DevExpress.XtraTreeList.StyleFormatConditions.StyleFormatCondition[] {
																																																													 styleFormatCondition1});
			this.tlsCuentas.ImageIndexFieldName = "";
			this.tlsCuentas.KeyFieldName = "idCuenta";
			this.tlsCuentas.Location = new System.Drawing.Point(8, 8);
			this.tlsCuentas.Name = "tlsCuentas";
			this.tlsCuentas.OptionsBehavior.UseTabKey = true;
			this.tlsCuentas.OptionsView.ShowSummaryFooter = true;
			this.tlsCuentas.ParentFieldName = "Padre";
			this.tlsCuentas.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																				this.repositoryItemSpinEdit1,
																																																				this.repositoryItemSpinEdit2,
																																																				this.repositoryItemSpinEdit3,
																																																				this.repositoryItemSpinEdit4,
																																																				this.repositoryItemTextEdit1,
																																																				this.repositoryItemTextEdit2,
																																																				this.repositoryItemCheckEdit1});
			this.tlsCuentas.RootValue = "0";
			this.tlsCuentas.Size = new System.Drawing.Size(1175, 407);
			this.tlsCuentas.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsCuentas.TabIndex = 227;
			this.tlsCuentas.DoubleClick += new System.EventHandler(this.tlsCuentas_DoubleClick);
			this.tlsCuentas.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.tlsCuentas_FocusedNodeChanged);
			// 
			// colCodigo
			// 
			this.colCodigo.Caption = "Código";
			this.colCodigo.ColumnEdit = this.repositoryItemTextEdit1;
			this.colCodigo.FieldName = "Codigo";
			this.colCodigo.MinWidth = 27;
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.OptionsColumn.AllowEdit = false;
			this.colCodigo.OptionsColumn.ReadOnly = true;
			this.colCodigo.VisibleIndex = 0;
			this.colCodigo.Width = 180;
			// 
			// repositoryItemTextEdit1
			// 
			this.repositoryItemTextEdit1.AutoHeight = false;
			this.repositoryItemTextEdit1.Name = "repositoryItemTextEdit1";
			// 
			// colDescripcion
			// 
			this.colDescripcion.Caption = "Descripción";
			this.colDescripcion.ColumnEdit = this.repositoryItemTextEdit2;
			this.colDescripcion.FieldName = "Descripcion";
			this.colDescripcion.Name = "colDescripcion";
			this.colDescripcion.OptionsColumn.AllowEdit = false;
			this.colDescripcion.OptionsColumn.ReadOnly = true;
			this.colDescripcion.VisibleIndex = 1;
			this.colDescripcion.Width = 319;
			// 
			// repositoryItemTextEdit2
			// 
			this.repositoryItemTextEdit2.AutoHeight = false;
			this.repositoryItemTextEdit2.Name = "repositoryItemTextEdit2";
			// 
			// colSaldoInicial
			// 
			this.colSaldoInicial.Caption = "Saldo Inicial";
			this.colSaldoInicial.ColumnEdit = this.repositoryItemSpinEdit1;
			this.colSaldoInicial.FieldName = "SaldoInicial";
			this.colSaldoInicial.Format.FormatString = "n2";
			this.colSaldoInicial.Name = "colSaldoInicial";
			this.colSaldoInicial.OptionsColumn.AllowEdit = false;
			this.colSaldoInicial.OptionsColumn.ReadOnly = true;
			this.colSaldoInicial.RowFooterSummaryStrFormat = "";
			this.colSaldoInicial.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.colSaldoInicial.SummaryFooterStrFormat = "";
			this.colSaldoInicial.VisibleIndex = 2;
			this.colSaldoInicial.Width = 140;
			// 
			// repositoryItemSpinEdit1
			// 
			this.repositoryItemSpinEdit1.AutoHeight = false;
			this.repositoryItemSpinEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit1.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit1.Name = "repositoryItemSpinEdit1";
			// 
			// ColDebe
			// 
			this.ColDebe.Caption = "Debe";
			this.ColDebe.ColumnEdit = this.repositoryItemSpinEdit2;
			this.ColDebe.FieldName = "Debe";
			this.ColDebe.Format.FormatString = "n2";
			this.ColDebe.Name = "ColDebe";
			this.ColDebe.OptionsColumn.AllowEdit = false;
			this.ColDebe.OptionsColumn.ReadOnly = true;
			this.ColDebe.RowFooterSummaryStrFormat = "";
			this.ColDebe.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.ColDebe.SummaryFooterStrFormat = "";
			this.ColDebe.VisibleIndex = 3;
			this.ColDebe.Width = 140;
			// 
			// repositoryItemSpinEdit2
			// 
			this.repositoryItemSpinEdit2.AutoHeight = false;
			this.repositoryItemSpinEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit2.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit2.Name = "repositoryItemSpinEdit2";
			// 
			// colHaber
			// 
			this.colHaber.Caption = "Haber";
			this.colHaber.ColumnEdit = this.repositoryItemSpinEdit3;
			this.colHaber.FieldName = "Haber";
			this.colHaber.Format.FormatString = "n2";
			this.colHaber.Name = "colHaber";
			this.colHaber.OptionsColumn.AllowEdit = false;
			this.colHaber.OptionsColumn.ReadOnly = true;
			this.colHaber.RowFooterSummaryStrFormat = "";
			this.colHaber.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.colHaber.SummaryFooterStrFormat = "";
			this.colHaber.VisibleIndex = 4;
			this.colHaber.Width = 140;
			// 
			// repositoryItemSpinEdit3
			// 
			this.repositoryItemSpinEdit3.AutoHeight = false;
			this.repositoryItemSpinEdit3.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit3.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit3.Name = "repositoryItemSpinEdit3";
			// 
			// colSaldoFinal
			// 
			this.colSaldoFinal.Caption = "Saldo Final";
			this.colSaldoFinal.ColumnEdit = this.repositoryItemSpinEdit4;
			this.colSaldoFinal.FieldName = "SaldoFinal";
			this.colSaldoFinal.Format.FormatString = "n2";
			this.colSaldoFinal.Name = "colSaldoFinal";
			this.colSaldoFinal.OptionsColumn.AllowEdit = false;
			this.colSaldoFinal.OptionsColumn.ReadOnly = true;
			this.colSaldoFinal.RowFooterSummaryStrFormat = "";
			this.colSaldoFinal.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.colSaldoFinal.SummaryFooterStrFormat = "";
			this.colSaldoFinal.VisibleIndex = 5;
			this.colSaldoFinal.Width = 140;
			// 
			// repositoryItemSpinEdit4
			// 
			this.repositoryItemSpinEdit4.AutoHeight = false;
			this.repositoryItemSpinEdit4.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit4.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit4.Name = "repositoryItemSpinEdit4";
			// 
			// Grupo
			// 
			this.Grupo.Caption = "Grupo";
			this.Grupo.ColumnEdit = this.repositoryItemCheckEdit1;
			this.Grupo.FieldName = "Grupo";
			this.Grupo.Name = "Grupo";
			this.Grupo.OptionsColumn.AllowEdit = false;
			this.Grupo.OptionsColumn.ReadOnly = true;
			this.Grupo.OptionsColumn.ShowInCustomizationForm = false;
			this.Grupo.VisibleIndex = 6;
			// 
			// repositoryItemCheckEdit1
			// 
			this.repositoryItemCheckEdit1.AutoHeight = false;
			this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(648, 11);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(29, 16);
			this.label12.TabIndex = 339;
			this.label12.Text = "Final";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(440, 11);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(33, 16);
			this.label13.TabIndex = 338;
			this.label13.Text = "Inicial";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBCFinal
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBCFinal.Appearance = appearance27;
			this.txtBCFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBCFinal.Enabled = false;
			this.txtBCFinal.Location = new System.Drawing.Point(704, 8);
			this.txtBCFinal.Name = "txtBCFinal";
			this.txtBCFinal.Size = new System.Drawing.Size(120, 21);
			this.txtBCFinal.TabIndex = 337;
			// 
			// txtBCInicial
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBCInicial.Appearance = appearance28;
			this.txtBCInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBCInicial.Location = new System.Drawing.Point(496, 8);
			this.txtBCInicial.Name = "txtBCInicial";
			this.txtBCInicial.Size = new System.Drawing.Size(120, 21);
			this.txtBCInicial.TabIndex = 336;
			this.txtBCInicial.TextChanged += new System.EventHandler(this.txtBCInicial_TextChanged);
			this.txtBCInicial.ValueChanged += new System.EventHandler(this.txtBCInicial_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 223;
			this.label1.Text = "Desde:";
			// 
			// dtBCDesde
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBCDesde.Appearance = appearance29;
			this.dtBCDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtBCDesde.DateButtons.Add(dateButton3);
			this.dtBCDesde.Format = "dd/MM/yyyy";
			this.dtBCDesde.Location = new System.Drawing.Point(72, 8);
			this.dtBCDesde.Name = "dtBCDesde";
			this.dtBCDesde.NonAutoSizeHeight = 23;
			this.dtBCDesde.Size = new System.Drawing.Size(112, 21);
			this.dtBCDesde.SpinButtonsVisible = true;
			this.dtBCDesde.TabIndex = 225;
			this.dtBCDesde.Value = ((object)(resources.GetObject("dtBCDesde.Value")));
			this.dtBCDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBCDesde_KeyDown);
			// 
			// dtBCHasta
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBCHasta.Appearance = appearance30;
			this.dtBCHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtBCHasta.DateButtons.Add(dateButton4);
			this.dtBCHasta.Format = "dd/MM/yyyy";
			this.dtBCHasta.Location = new System.Drawing.Point(288, 8);
			this.dtBCHasta.Name = "dtBCHasta";
			this.dtBCHasta.NonAutoSizeHeight = 23;
			this.dtBCHasta.Size = new System.Drawing.Size(112, 21);
			this.dtBCHasta.SpinButtonsVisible = true;
			this.dtBCHasta.TabIndex = 226;
			this.dtBCHasta.Value = ((object)(resources.GetObject("dtBCHasta.Value")));
			this.dtBCHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBCHasta_KeyDown);
			// 
			// btnGeneraBalance
			// 
			this.btnGeneraBalance.Location = new System.Drawing.Point(848, 8);
			this.btnGeneraBalance.Name = "btnGeneraBalance";
			this.btnGeneraBalance.Size = new System.Drawing.Size(80, 21);
			this.btnGeneraBalance.TabIndex = 218;
			this.btnGeneraBalance.Text = "Generar";
			this.btnGeneraBalance.Click += new System.EventHandler(this.btnGeneraBalance_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(224, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 224;
			this.label2.Text = "Hasta:";
			// 
			// btnImprimirBC
			// 
			this.btnImprimirBC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimirBC.Location = new System.Drawing.Point(1047, 8);
			this.btnImprimirBC.Name = "btnImprimirBC";
			this.btnImprimirBC.Size = new System.Drawing.Size(72, 21);
			this.btnImprimirBC.TabIndex = 220;
			this.btnImprimirBC.Text = "&Imprimir";
			this.btnImprimirBC.Click += new System.EventHandler(this.btnImprimirBC_Click);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.tlsBalanceGeneral);
			this.ultraTabPageControl2.Controls.Add(this.btnGeneraBalanceGeneral);
			this.ultraTabPageControl2.Controls.Add(this.label8);
			this.ultraTabPageControl2.Controls.Add(this.dtBGDesde);
			this.ultraTabPageControl2.Controls.Add(this.dtBGHasta);
			this.ultraTabPageControl2.Controls.Add(this.label9);
			this.ultraTabPageControl2.Controls.Add(this.btnImprimirBG);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1189, 418);
			// 
			// tlsBalanceGeneral
			// 
			this.tlsBalanceGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsBalanceGeneral.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.tlsBalanceGeneral.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																																		 this.treeListColumn1,
																																																		 this.treeListColumn2,
																																																		 this.treeListColumn3,
																																																		 this.treeListColumn4,
																																																		 this.treeListColumn5,
																																																		 this.treeListColumn6});
			this.tlsBalanceGeneral.ImageIndexFieldName = "";
			this.tlsBalanceGeneral.KeyFieldName = "idCuenta";
			this.tlsBalanceGeneral.Location = new System.Drawing.Point(7, 8);
			this.tlsBalanceGeneral.Name = "tlsBalanceGeneral";
			this.tlsBalanceGeneral.OptionsBehavior.UseTabKey = true;
			this.tlsBalanceGeneral.OptionsView.ShowSummaryFooter = true;
			this.tlsBalanceGeneral.ParentFieldName = "Padre";
			this.tlsBalanceGeneral.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																							 this.repositoryItemSpinEdit5,
																																																							 this.repositoryItemSpinEdit6,
																																																							 this.repositoryItemSpinEdit7,
																																																							 this.repositoryItemSpinEdit8,
																																																							 this.repositoryItemTextEdit3,
																																																							 this.repositoryItemTextEdit4});
			this.tlsBalanceGeneral.RootValue = "0";
			this.tlsBalanceGeneral.Size = new System.Drawing.Size(1175, 407);
			this.tlsBalanceGeneral.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsBalanceGeneral.TabIndex = 235;
			// 
			// treeListColumn1
			// 
			this.treeListColumn1.Caption = "Código";
			this.treeListColumn1.ColumnEdit = this.repositoryItemTextEdit3;
			this.treeListColumn1.FieldName = "Codigo";
			this.treeListColumn1.MinWidth = 27;
			this.treeListColumn1.Name = "treeListColumn1";
			this.treeListColumn1.OptionsColumn.AllowEdit = false;
			this.treeListColumn1.OptionsColumn.ReadOnly = true;
			this.treeListColumn1.VisibleIndex = 0;
			this.treeListColumn1.Width = 180;
			// 
			// repositoryItemTextEdit3
			// 
			this.repositoryItemTextEdit3.AutoHeight = false;
			this.repositoryItemTextEdit3.Name = "repositoryItemTextEdit3";
			// 
			// treeListColumn2
			// 
			this.treeListColumn2.Caption = "Descripción";
			this.treeListColumn2.ColumnEdit = this.repositoryItemTextEdit4;
			this.treeListColumn2.FieldName = "Descripcion";
			this.treeListColumn2.Name = "treeListColumn2";
			this.treeListColumn2.OptionsColumn.AllowEdit = false;
			this.treeListColumn2.OptionsColumn.ReadOnly = true;
			this.treeListColumn2.VisibleIndex = 1;
			this.treeListColumn2.Width = 319;
			// 
			// repositoryItemTextEdit4
			// 
			this.repositoryItemTextEdit4.AutoHeight = false;
			this.repositoryItemTextEdit4.Name = "repositoryItemTextEdit4";
			// 
			// treeListColumn3
			// 
			this.treeListColumn3.Caption = "Saldo Inicial";
			this.treeListColumn3.ColumnEdit = this.repositoryItemSpinEdit5;
			this.treeListColumn3.FieldName = "SaldoInicial";
			this.treeListColumn3.Format.FormatString = "n2";
			this.treeListColumn3.Name = "treeListColumn3";
			this.treeListColumn3.OptionsColumn.AllowEdit = false;
			this.treeListColumn3.OptionsColumn.ReadOnly = true;
			this.treeListColumn3.RowFooterSummaryStrFormat = "";
			this.treeListColumn3.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn3.SummaryFooterStrFormat = "";
			this.treeListColumn3.VisibleIndex = 2;
			this.treeListColumn3.Width = 140;
			// 
			// repositoryItemSpinEdit5
			// 
			this.repositoryItemSpinEdit5.AutoHeight = false;
			this.repositoryItemSpinEdit5.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit5.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit5.Name = "repositoryItemSpinEdit5";
			// 
			// treeListColumn4
			// 
			this.treeListColumn4.Caption = "Debe";
			this.treeListColumn4.ColumnEdit = this.repositoryItemSpinEdit6;
			this.treeListColumn4.FieldName = "Debe";
			this.treeListColumn4.Format.FormatString = "n2";
			this.treeListColumn4.Name = "treeListColumn4";
			this.treeListColumn4.OptionsColumn.AllowEdit = false;
			this.treeListColumn4.OptionsColumn.ReadOnly = true;
			this.treeListColumn4.RowFooterSummaryStrFormat = "";
			this.treeListColumn4.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn4.SummaryFooterStrFormat = "";
			this.treeListColumn4.VisibleIndex = 3;
			this.treeListColumn4.Width = 140;
			// 
			// repositoryItemSpinEdit6
			// 
			this.repositoryItemSpinEdit6.AutoHeight = false;
			this.repositoryItemSpinEdit6.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit6.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit6.Name = "repositoryItemSpinEdit6";
			// 
			// treeListColumn5
			// 
			this.treeListColumn5.Caption = "Haber";
			this.treeListColumn5.ColumnEdit = this.repositoryItemSpinEdit7;
			this.treeListColumn5.FieldName = "Haber";
			this.treeListColumn5.Format.FormatString = "n2";
			this.treeListColumn5.Name = "treeListColumn5";
			this.treeListColumn5.OptionsColumn.AllowEdit = false;
			this.treeListColumn5.OptionsColumn.ReadOnly = true;
			this.treeListColumn5.RowFooterSummaryStrFormat = "";
			this.treeListColumn5.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn5.SummaryFooterStrFormat = "";
			this.treeListColumn5.VisibleIndex = 4;
			this.treeListColumn5.Width = 140;
			// 
			// repositoryItemSpinEdit7
			// 
			this.repositoryItemSpinEdit7.AutoHeight = false;
			this.repositoryItemSpinEdit7.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit7.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit7.Name = "repositoryItemSpinEdit7";
			// 
			// treeListColumn6
			// 
			this.treeListColumn6.Caption = "Saldo Final";
			this.treeListColumn6.ColumnEdit = this.repositoryItemSpinEdit8;
			this.treeListColumn6.FieldName = "SaldoFinal";
			this.treeListColumn6.Format.FormatString = "n2";
			this.treeListColumn6.Name = "treeListColumn6";
			this.treeListColumn6.OptionsColumn.AllowEdit = false;
			this.treeListColumn6.OptionsColumn.ReadOnly = true;
			this.treeListColumn6.RowFooterSummaryStrFormat = "";
			this.treeListColumn6.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn6.SummaryFooterStrFormat = "";
			this.treeListColumn6.VisibleIndex = 5;
			this.treeListColumn6.Width = 140;
			// 
			// repositoryItemSpinEdit8
			// 
			this.repositoryItemSpinEdit8.AutoHeight = false;
			this.repositoryItemSpinEdit8.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit8.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit8.Name = "repositoryItemSpinEdit8";
			// 
			// btnGeneraBalanceGeneral
			// 
			this.btnGeneraBalanceGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGeneraBalanceGeneral.Location = new System.Drawing.Point(511, 8);
			this.btnGeneraBalanceGeneral.Name = "btnGeneraBalanceGeneral";
			this.btnGeneraBalanceGeneral.Size = new System.Drawing.Size(72, 21);
			this.btnGeneraBalanceGeneral.TabIndex = 234;
			this.btnGeneraBalanceGeneral.Text = "&Generar";
			this.btnGeneraBalanceGeneral.Click += new System.EventHandler(this.btnGeneraBalanceGeneral_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 11);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 16);
			this.label8.TabIndex = 228;
			this.label8.Text = "Desde:";
			// 
			// dtBGDesde
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBGDesde.Appearance = appearance31;
			this.dtBGDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtBGDesde.DateButtons.Add(dateButton5);
			this.dtBGDesde.Format = "dd/MM/yyyy";
			this.dtBGDesde.Location = new System.Drawing.Point(72, 8);
			this.dtBGDesde.Name = "dtBGDesde";
			this.dtBGDesde.NonAutoSizeHeight = 23;
			this.dtBGDesde.Size = new System.Drawing.Size(112, 21);
			this.dtBGDesde.SpinButtonsVisible = true;
			this.dtBGDesde.TabIndex = 230;
			this.dtBGDesde.Value = ((object)(resources.GetObject("dtBGDesde.Value")));
			this.dtBGDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBGDesde_KeyDown);
			// 
			// dtBGHasta
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBGHasta.Appearance = appearance32;
			this.dtBGHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtBGHasta.DateButtons.Add(dateButton6);
			this.dtBGHasta.Format = "dd/MM/yyyy";
			this.dtBGHasta.Location = new System.Drawing.Point(288, 8);
			this.dtBGHasta.Name = "dtBGHasta";
			this.dtBGHasta.NonAutoSizeHeight = 23;
			this.dtBGHasta.Size = new System.Drawing.Size(112, 21);
			this.dtBGHasta.SpinButtonsVisible = true;
			this.dtBGHasta.TabIndex = 231;
			this.dtBGHasta.Value = ((object)(resources.GetObject("dtBGHasta.Value")));
			this.dtBGHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBGHasta_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(224, 11);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(37, 16);
			this.label9.TabIndex = 229;
			this.label9.Text = "Hasta:";
			// 
			// btnImprimirBG
			// 
			this.btnImprimirBG.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimirBG.Location = new System.Drawing.Point(671, 8);
			this.btnImprimirBG.Name = "btnImprimirBG";
			this.btnImprimirBG.Size = new System.Drawing.Size(72, 21);
			this.btnImprimirBG.TabIndex = 227;
			this.btnImprimirBG.Text = "&Imprimir";
			this.btnImprimirBG.Click += new System.EventHandler(this.btnImprimirBG_Click);
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.tlsBalanceDeResultados);
			this.pgEntrega.Controls.Add(this.btnGenerarBR);
			this.pgEntrega.Controls.Add(this.label10);
			this.pgEntrega.Controls.Add(this.dtBRDesde);
			this.pgEntrega.Controls.Add(this.dtBRHasta);
			this.pgEntrega.Controls.Add(this.label11);
			this.pgEntrega.Controls.Add(this.btnImprimirBR);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(1189, 418);
			// 
			// tlsBalanceDeResultados
			// 
			this.tlsBalanceDeResultados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsBalanceDeResultados.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.tlsBalanceDeResultados.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																																					this.treeListColumn7,
																																																					this.treeListColumn8,
																																																					this.treeListColumn9,
																																																					this.treeListColumn10,
																																																					this.treeListColumn11,
																																																					this.treeListColumn12});
			this.tlsBalanceDeResultados.ImageIndexFieldName = "";
			this.tlsBalanceDeResultados.KeyFieldName = "idCuenta";
			this.tlsBalanceDeResultados.Location = new System.Drawing.Point(7, 8);
			this.tlsBalanceDeResultados.Name = "tlsBalanceDeResultados";
			this.tlsBalanceDeResultados.OptionsBehavior.UseTabKey = true;
			this.tlsBalanceDeResultados.OptionsView.ShowSummaryFooter = true;
			this.tlsBalanceDeResultados.ParentFieldName = "Padre";
			this.tlsBalanceDeResultados.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																										this.repositoryItemSpinEdit9,
																																																										this.repositoryItemSpinEdit10,
																																																										this.repositoryItemSpinEdit11,
																																																										this.repositoryItemSpinEdit12,
																																																										this.repositoryItemTextEdit5,
																																																										this.repositoryItemTextEdit6});
			this.tlsBalanceDeResultados.RootValue = "0";
			this.tlsBalanceDeResultados.Size = new System.Drawing.Size(1175, 407);
			this.tlsBalanceDeResultados.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsBalanceDeResultados.TabIndex = 237;
			// 
			// treeListColumn7
			// 
			this.treeListColumn7.Caption = "Código";
			this.treeListColumn7.ColumnEdit = this.repositoryItemTextEdit5;
			this.treeListColumn7.FieldName = "Codigo";
			this.treeListColumn7.MinWidth = 27;
			this.treeListColumn7.Name = "treeListColumn7";
			this.treeListColumn7.OptionsColumn.AllowEdit = false;
			this.treeListColumn7.OptionsColumn.ReadOnly = true;
			this.treeListColumn7.VisibleIndex = 0;
			this.treeListColumn7.Width = 180;
			// 
			// repositoryItemTextEdit5
			// 
			this.repositoryItemTextEdit5.AutoHeight = false;
			this.repositoryItemTextEdit5.Name = "repositoryItemTextEdit5";
			// 
			// treeListColumn8
			// 
			this.treeListColumn8.Caption = "Descripción";
			this.treeListColumn8.ColumnEdit = this.repositoryItemTextEdit6;
			this.treeListColumn8.FieldName = "Descripcion";
			this.treeListColumn8.Name = "treeListColumn8";
			this.treeListColumn8.OptionsColumn.AllowEdit = false;
			this.treeListColumn8.OptionsColumn.ReadOnly = true;
			this.treeListColumn8.VisibleIndex = 1;
			this.treeListColumn8.Width = 319;
			// 
			// repositoryItemTextEdit6
			// 
			this.repositoryItemTextEdit6.AutoHeight = false;
			this.repositoryItemTextEdit6.Name = "repositoryItemTextEdit6";
			// 
			// treeListColumn9
			// 
			this.treeListColumn9.Caption = "Saldo Inicial";
			this.treeListColumn9.ColumnEdit = this.repositoryItemSpinEdit9;
			this.treeListColumn9.FieldName = "SaldoInicial";
			this.treeListColumn9.Format.FormatString = "n2";
			this.treeListColumn9.Name = "treeListColumn9";
			this.treeListColumn9.OptionsColumn.AllowEdit = false;
			this.treeListColumn9.OptionsColumn.ReadOnly = true;
			this.treeListColumn9.RowFooterSummaryStrFormat = "";
			this.treeListColumn9.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn9.SummaryFooterStrFormat = "";
			this.treeListColumn9.VisibleIndex = 2;
			this.treeListColumn9.Width = 140;
			// 
			// repositoryItemSpinEdit9
			// 
			this.repositoryItemSpinEdit9.AutoHeight = false;
			this.repositoryItemSpinEdit9.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit9.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit9.Name = "repositoryItemSpinEdit9";
			// 
			// treeListColumn10
			// 
			this.treeListColumn10.Caption = "Debe";
			this.treeListColumn10.ColumnEdit = this.repositoryItemSpinEdit10;
			this.treeListColumn10.FieldName = "Debe";
			this.treeListColumn10.Format.FormatString = "n2";
			this.treeListColumn10.Name = "treeListColumn10";
			this.treeListColumn10.OptionsColumn.AllowEdit = false;
			this.treeListColumn10.OptionsColumn.ReadOnly = true;
			this.treeListColumn10.RowFooterSummaryStrFormat = "";
			this.treeListColumn10.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn10.SummaryFooterStrFormat = "";
			this.treeListColumn10.VisibleIndex = 3;
			this.treeListColumn10.Width = 140;
			// 
			// repositoryItemSpinEdit10
			// 
			this.repositoryItemSpinEdit10.AutoHeight = false;
			this.repositoryItemSpinEdit10.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit10.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit10.Name = "repositoryItemSpinEdit10";
			// 
			// treeListColumn11
			// 
			this.treeListColumn11.Caption = "Haber";
			this.treeListColumn11.ColumnEdit = this.repositoryItemSpinEdit11;
			this.treeListColumn11.FieldName = "Haber";
			this.treeListColumn11.Format.FormatString = "n2";
			this.treeListColumn11.Name = "treeListColumn11";
			this.treeListColumn11.OptionsColumn.AllowEdit = false;
			this.treeListColumn11.OptionsColumn.ReadOnly = true;
			this.treeListColumn11.RowFooterSummaryStrFormat = "";
			this.treeListColumn11.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn11.SummaryFooterStrFormat = "";
			this.treeListColumn11.VisibleIndex = 4;
			this.treeListColumn11.Width = 140;
			// 
			// repositoryItemSpinEdit11
			// 
			this.repositoryItemSpinEdit11.AutoHeight = false;
			this.repositoryItemSpinEdit11.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit11.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit11.Name = "repositoryItemSpinEdit11";
			// 
			// treeListColumn12
			// 
			this.treeListColumn12.Caption = "Saldo Final";
			this.treeListColumn12.ColumnEdit = this.repositoryItemSpinEdit12;
			this.treeListColumn12.FieldName = "SaldoFinal";
			this.treeListColumn12.Format.FormatString = "n2";
			this.treeListColumn12.Name = "treeListColumn12";
			this.treeListColumn12.OptionsColumn.AllowEdit = false;
			this.treeListColumn12.OptionsColumn.ReadOnly = true;
			this.treeListColumn12.RowFooterSummaryStrFormat = "";
			this.treeListColumn12.SummaryFooter = DevExpress.XtraTreeList.SummaryItemType.Sum;
			this.treeListColumn12.SummaryFooterStrFormat = "";
			this.treeListColumn12.VisibleIndex = 5;
			this.treeListColumn12.Width = 140;
			// 
			// repositoryItemSpinEdit12
			// 
			this.repositoryItemSpinEdit12.AutoHeight = false;
			this.repositoryItemSpinEdit12.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton()});
			this.repositoryItemSpinEdit12.DisplayFormat.FormatString = "n2";
			this.repositoryItemSpinEdit12.Name = "repositoryItemSpinEdit12";
			// 
			// btnGenerarBR
			// 
			this.btnGenerarBR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerarBR.Location = new System.Drawing.Point(527, 8);
			this.btnGenerarBR.Name = "btnGenerarBR";
			this.btnGenerarBR.Size = new System.Drawing.Size(72, 21);
			this.btnGenerarBR.TabIndex = 238;
			this.btnGenerarBR.Text = "&Generar";
			this.btnGenerarBR.Click += new System.EventHandler(this.btnGenerarBR_Click);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(8, 11);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(40, 16);
			this.label10.TabIndex = 233;
			this.label10.Text = "Desde:";
			// 
			// dtBRDesde
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBRDesde.Appearance = appearance33;
			this.dtBRDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton7.Caption = "Today";
			this.dtBRDesde.DateButtons.Add(dateButton7);
			this.dtBRDesde.Format = "dd/MM/yyyy";
			this.dtBRDesde.Location = new System.Drawing.Point(72, 8);
			this.dtBRDesde.Name = "dtBRDesde";
			this.dtBRDesde.NonAutoSizeHeight = 23;
			this.dtBRDesde.Size = new System.Drawing.Size(112, 21);
			this.dtBRDesde.SpinButtonsVisible = true;
			this.dtBRDesde.TabIndex = 235;
			this.dtBRDesde.Value = ((object)(resources.GetObject("dtBRDesde.Value")));
			this.dtBRDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBRDesde_KeyDown);
			// 
			// dtBRHasta
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtBRHasta.Appearance = appearance34;
			this.dtBRHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton8.Caption = "Today";
			this.dtBRHasta.DateButtons.Add(dateButton8);
			this.dtBRHasta.Format = "dd/MM/yyyy";
			this.dtBRHasta.Location = new System.Drawing.Point(288, 8);
			this.dtBRHasta.Name = "dtBRHasta";
			this.dtBRHasta.NonAutoSizeHeight = 23;
			this.dtBRHasta.Size = new System.Drawing.Size(112, 21);
			this.dtBRHasta.SpinButtonsVisible = true;
			this.dtBRHasta.TabIndex = 236;
			this.dtBRHasta.Value = ((object)(resources.GetObject("dtBRHasta.Value")));
			this.dtBRHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtBRHasta_KeyDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(224, 11);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 234;
			this.label11.Text = "Hasta:";
			// 
			// btnImprimirBR
			// 
			this.btnImprimirBR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimirBR.Location = new System.Drawing.Point(615, 8);
			this.btnImprimirBR.Name = "btnImprimirBR";
			this.btnImprimirBR.Size = new System.Drawing.Size(72, 21);
			this.btnImprimirBR.TabIndex = 232;
			this.btnImprimirBR.Text = "&Imprimir";
			this.btnImprimirBR.Click += new System.EventHandler(this.btnImprimirBR_Click);
			// 
			// utpMayorGeneralCondensado
			// 
			this.utpMayorGeneralCondensado.Controls.Add(this.uGridMayorCondensado);
			this.utpMayorGeneralCondensado.Location = new System.Drawing.Point(-10000, -10000);
			this.utpMayorGeneralCondensado.Name = "utpMayorGeneralCondensado";
			this.utpMayorGeneralCondensado.Size = new System.Drawing.Size(1189, 418);
			// 
			// uGridMayorCondensado
			// 
			this.uGridMayorCondensado.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridMayorCondensado.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMayorCondensado.DataSource = this.ultraDataSource3;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridMayorCondensado.DisplayLayout.Appearance = appearance35;
			this.uGridMayorCondensado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 62;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 237;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 429;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance36;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Saldo Inicial";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 104;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance37;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 128;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance38;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 128;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance39;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Width = 128;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			ultraGridBand2.Header.Caption = "";
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand2.Override.GroupByColumnAppearance = appearance40;
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand2.Override.GroupByColumnHeaderAppearance = appearance41;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance42;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.Hidden = true;
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance43;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance44;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridMayorCondensado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridMayorCondensado.DisplayLayout.GroupByBox.Hidden = true;
			appearance45.ForeColor = System.Drawing.Color.Black;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMayorCondensado.DisplayLayout.Override.ActiveRowAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.Transparent;
			this.uGridMayorCondensado.DisplayLayout.Override.CardAreaAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance47.FontData.BoldAsString = "True";
			appearance47.FontData.Name = "Arial";
			appearance47.FontData.SizeInPoints = 8F;
			appearance47.ForeColor = System.Drawing.Color.White;
			appearance47.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMayorCondensado.DisplayLayout.Override.HeaderAppearance = appearance47;
			appearance48.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance48.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMayorCondensado.DisplayLayout.Override.RowAlternateAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMayorCondensado.DisplayLayout.Override.RowSelectorAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMayorCondensado.DisplayLayout.Override.SelectedRowAppearance = appearance50;
			this.uGridMayorCondensado.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridMayorCondensado.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMayorCondensado.Location = new System.Drawing.Point(7, 8);
			this.uGridMayorCondensado.Name = "uGridMayorCondensado";
			this.uGridMayorCondensado.Size = new System.Drawing.Size(1175, 401);
			this.uGridMayorCondensado.TabIndex = 9;
			this.uGridMayorCondensado.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridMayorCondensado_InitializeRow);
			this.uGridMayorCondensado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridMayorCondensado_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Double);
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn21.DataType = typeof(System.Double);
			ultraDataColumn22.DataType = typeof(System.Double);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// btExportar
			// 
			this.btExportar.Location = new System.Drawing.Point(184, 40);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(80, 21);
			this.btExportar.TabIndex = 629;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Location = new System.Drawing.Point(96, 40);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 21);
			this.btnImprimir.TabIndex = 338;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Location = new System.Drawing.Point(272, 40);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 21);
			this.btnAsiento.TabIndex = 337;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(bool);
			ultraDataColumn24.DefaultValue = false;
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn27.DefaultValue = 0;
			ultraDataColumn28.DataType = typeof(System.Double);
			ultraDataColumn28.DefaultValue = 0;
			ultraDataColumn29.DataType = typeof(System.Double);
			ultraDataColumn29.DefaultValue = 0;
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn30.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24,
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
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1189, 418);
			// 
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance51;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.pgEntrega);
			this.utcReportes.Controls.Add(this.ultraTabPageControl9);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Controls.Add(this.ultraTabPageControl2);
			this.utcReportes.Controls.Add(this.utpMayorGeneralCondensado);
			this.utcReportes.Location = new System.Drawing.Point(8, 72);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(1192, 440);
			this.utcReportes.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.PropertyPage2003;
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.OfficeXPToolbarButton;
			this.utcReportes.TabIndex = 228;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Mayor General";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Balance De Comprobacion";
			ultraTab3.TabPage = this.ultraTabPageControl2;
			ultraTab3.Text = "Balance General";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Balance de Resultados";
			ultraTab5.TabPage = this.utpMayorGeneralCondensado;
			ultraTab5.Text = "Mayor General Condensado";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3,
																																													ultraTab4,
																																													ultraTab5});
			this.utcReportes.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcReportes.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcReportes_SelectedTabChanged);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.FileLimitBehaviour = Infragistics.Win.UltraWinGrid.ExcelExport.FileLimitBehaviour.TruncateData;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(784, 11);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 16);
			this.label14.TabIndex = 640;
			this.label14.Text = "Cuenta";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuentas
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentas.Appearance = appearance52;
			this.txtCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentas.Location = new System.Drawing.Point(848, 8);
			this.txtCuentas.Name = "txtCuentas";
			this.txtCuentas.Size = new System.Drawing.Size(216, 21);
			this.txtCuentas.TabIndex = 634;
			this.txtCuentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuentas_KeyDown);
			this.txtCuentas.ValueChanged += new System.EventHandler(this.txtCuentas_ValueChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(592, 11);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(29, 16);
			this.label15.TabIndex = 639;
			this.label15.Text = "Final";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(392, 11);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(33, 16);
			this.label16.TabIndex = 638;
			this.label16.Text = "Inicial";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuentaFinal
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaFinal.Appearance = appearance53;
			this.txtCuentaFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaFinal.Enabled = false;
			this.txtCuentaFinal.Location = new System.Drawing.Point(640, 8);
			this.txtCuentaFinal.Name = "txtCuentaFinal";
			this.txtCuentaFinal.Size = new System.Drawing.Size(120, 21);
			this.txtCuentaFinal.TabIndex = 633;
			this.txtCuentaFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuentaFinal_KeyDown);
			this.txtCuentaFinal.ValueChanged += new System.EventHandler(this.txtCuentaFinal_ValueChanged);
			// 
			// txtCuentaInicial
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaInicial.Appearance = appearance54;
			this.txtCuentaInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaInicial.Location = new System.Drawing.Point(448, 8);
			this.txtCuentaInicial.Name = "txtCuentaInicial";
			this.txtCuentaInicial.Size = new System.Drawing.Size(120, 21);
			this.txtCuentaInicial.TabIndex = 632;
			this.txtCuentaInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuentaInicial_KeyDown);
			this.txtCuentaInicial.TextChanged += new System.EventHandler(this.txtCuentaInicial_TextChanged);
			this.txtCuentaInicial.ValueChanged += new System.EventHandler(this.txtCuentaInicial_ValueChanged);
			// 
			// dtHasta
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance55;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton9.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton9);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 631;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance56;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton10.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton10);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 630;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 11);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 637;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 11);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 636;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Location = new System.Drawing.Point(8, 40);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(80, 21);
			this.btnGenerar.TabIndex = 635;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(880, 32);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 641;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(368, 43);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(84, 16);
			this.label19.TabIndex = 643;
			this.label19.Text = "Centro de costo";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbProyecto
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance57;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 408;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(464, 40);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(248, 21);
			this.cmbProyecto.TabIndex = 642;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// btnVacio
			// 
			this.btnVacio.Location = new System.Drawing.Point(728, 40);
			this.btnVacio.Name = "btnVacio";
			this.btnVacio.Size = new System.Drawing.Size(32, 21);
			this.btnVacio.TabIndex = 644;
			this.btnVacio.Text = "...";
			this.btnVacio.Click += new System.EventHandler(this.btnVacio_Click);
			// 
			// frmReportesDeContabilidad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1208, 518);
			this.Controls.Add(this.btnVacio);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtCuentas);
			this.Controls.Add(this.txtCuentaFinal);
			this.Controls.Add(this.txtCuentaInicial);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.utcReportes);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btExportar);
			this.KeyPreview = true;
			this.Name = "frmReportesDeContabilidad";
			this.Text = "Reportes de Contabilidad";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportesDeContabilidad_KeyDown);
			this.Load += new System.EventHandler(this.frmReportesDeContabilidad_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalDebe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMayorHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMayorDesde)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBCFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBCInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBCDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBCHasta)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.tlsBalanceGeneral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBGDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBGHasta)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.tlsBalanceDeResultados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemSpinEdit12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBRDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtBRHasta)).EndInit();
			this.utpMayorGeneralCondensado.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridMayorCondensado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void btConsultar_Click1(object sender, System.EventArgs e)
		{
			
		}

		private void frmReportesDeContabilidad_Load(object sender, System.EventArgs e)
		{			
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtMayorDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtMayorHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtBCDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtBCHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtBGDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtBGHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtBRDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtBRHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order By Nombre");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void tlsCuentas_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void tlsCuentas_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{

		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				int idAsiento = 0;
				idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Asiento Where idAsiento = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value));

				frmAsientos Asientos = new frmAsientos(((int)this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value));
				Asientos.MdiParent = MdiParent;
				Asientos.Show();
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			if (this.txtCuentaInicial.Text.ToString().Length > 0 && this.txtCuentaFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuentaFinal.Focus();
				return;	
			}

			if (this.txtCuentaInicial.Text.ToString().Length == 0 && this.txtCuentaFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuentaInicial.Focus();
				return;
			}
		
			bool bRango = false;

			if (this.txtCuentaInicial.Text.ToString().Length > 0 && this.txtCuentaFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtCuentaInicial.Text.ToString();
				string sFinal = this.txtCuentaFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (Int64.Parse(sInicial) > Int64.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCuentaInicial.Focus();
					return;
				}

				bRango = true;
			}
			
			bool bCta = false;

			if (this.txtCuenta.Text.ToString().Length > 0) bCta = true;
			
			string sReporte = "ReporteMayorGeneral.rpt";
			string sTitulo = "Mayor General";

			//int idEmpresa = 0;
			//if (MenuLatinium.stDirFacturacion == "CROMATIKA") idEmpresa = 4;
			int idEmpresa = MenuLatinium.stIdDB;

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

//			if (this.utcReportes.Tabs[0].Selected) 
//			{
//				ParameterField paramFieldRango = new ParameterField ();
//				ParameterDiscreteValue discreteValRango = new ParameterDiscreteValue ();
//				paramFieldRango.ParameterFieldName = "@Rango";
//				discreteValRango.Value = bRango; 
//				paramFieldRango.CurrentValues.Add (discreteValRango);
//				paramFields.Add (paramFieldRango);
//
//				ParameterField paramFieldCuenta = new ParameterField ();
//				ParameterDiscreteValue discreteValCuenta = new ParameterDiscreteValue ();
//				paramFieldCuenta.ParameterFieldName = "@Cuenta";
//				discreteValCuenta.Value = this.txtCuentas.Text.ToString(); 
//				paramFieldCuenta.CurrentValues.Add (discreteValCuenta);
//				paramFields.Add (paramFieldCuenta);
//
//				ParameterField paramFieldCta = new ParameterField ();
//				ParameterDiscreteValue discreteValCta = new ParameterDiscreteValue ();
//				paramFieldCta.ParameterFieldName = "@Cta";
//				discreteValCta.Value = bCta; 
//				paramFieldCta.CurrentValues.Add (discreteValCta);
//				paramFields.Add (paramFieldCta);			
//			}
//			else 
			if (this.utcReportes.Tabs[1].Selected) 
			{
				ParameterField paramFieldInicial = new ParameterField ();
				ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
				paramFieldInicial.ParameterFieldName = "@Inicial";
				discreteValInicial.Value = this.txtCuentaInicial.Text.ToString(); 
				paramFieldInicial.CurrentValues.Add (discreteValInicial);
				paramFields.Add (paramFieldInicial);

				ParameterField paramFieldFinal = new ParameterField ();
				ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
				paramFieldFinal.ParameterFieldName = "@Final";
				discreteValFinal.Value = this.txtCuentaFinal.Text.ToString(); 
				paramFieldFinal.CurrentValues.Add (discreteValFinal);
				paramFields.Add (paramFieldFinal);

				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				sReporte = "BalanceDeComprobacion4.rpt";
				sTitulo = "Balance De Comprobación";				
			}
			else if (this.utcReportes.Tabs[2].Selected) 
			{
				ParameterField paramFieldInicial = new ParameterField ();
				ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
				paramFieldInicial.ParameterFieldName = "@Inicial";
				discreteValInicial.Value = this.txtCuentaInicial.Text.ToString(); 
				paramFieldInicial.CurrentValues.Add (discreteValInicial);
				paramFields.Add (paramFieldInicial);

				ParameterField paramFieldFinal = new ParameterField ();
				ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
				paramFieldFinal.ParameterFieldName = "@Final";
				discreteValFinal.Value = this.txtCuentaFinal.Text.ToString(); 
				paramFieldFinal.CurrentValues.Add (discreteValFinal);
				paramFields.Add (paramFieldFinal);

				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				sReporte = "BalanceGeneral.rpt";
				sTitulo = "Balance General";
			}
			else if (this.utcReportes.Tabs[3].Selected) 
			{
				ParameterField paramFieldInicial = new ParameterField ();
				ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
				paramFieldInicial.ParameterFieldName = "@Inicial";
				discreteValInicial.Value = this.txtCuentaInicial.Text.ToString(); 
				paramFieldInicial.CurrentValues.Add (discreteValInicial);
				paramFields.Add (paramFieldInicial);

				ParameterField paramFieldFinal = new ParameterField ();
				ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
				paramFieldFinal.ParameterFieldName = "@Final";
				discreteValFinal.Value = this.txtCuentaFinal.Text.ToString(); 
				paramFieldFinal.CurrentValues.Add (discreteValFinal);
				paramFields.Add (paramFieldFinal);

				ParameterField paramFieldEmpresa = new ParameterField ();
				ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
				paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
				discreteValEmpresa.Value = idEmpresa; 
				paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
				paramFields.Add (paramFieldEmpresa);

				sReporte = "BalanceDeResultados.rpt";
				sTitulo = "Balance De Resultados";				
			}
			
			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo(sTitulo);
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void dtMayorDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtMayorHasta.Focus();
		}

		private void dtMayorHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtInicial.Focus();
		}

		private void txtInicial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtFinal.Enabled)this.txtFinal.Focus();
				else this.txtCuenta.Focus();
			}
		}

		private void txtInicial_TextChanged(object sender, System.EventArgs e)
		{
			if (this.txtInicial.Text.Length > 0) 
			{
				this.txtFinal.Enabled = true;
				this.txtCuenta.Enabled = false;
			}
			else if (this.txtInicial.Text.Length == 0) 
			{
				this.txtFinal.Enabled = false;
				this.txtCuenta.Enabled = true;
				this.txtFinal.Text = "";
			}
		}

		private void txtFinal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCuenta.Enabled)this.txtCuenta.Focus();
				else this.btConsultar_Click(sender, e);
			}
		}

		private void txtCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btConsultar_Click(sender, e);
		}

		private void btConsultar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtMayorDesde.Value;
			DateTime dtHasta = (DateTime)this.dtMayorHasta.Value;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFinal.Focus();
				return;
			}

			if (this.txtInicial.Text.ToString().Length == 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtInicial.Focus();
				return;
			}
		
			bool bRango = false;

			if (this.txtInicial.Text.ToString().Length > 0 && this.txtFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtInicial.Text.ToString();
				string sFinal = this.txtFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (int.Parse(sInicial) > int.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtInicial.Focus();
					return;
				}

				bRango = true;
			}
			
			bool bCta = false;

			if (this.txtCuenta.Text.ToString().Length > 0) bCta = true;

			this.Cursor = Cursors.WaitCursor;

			this.ultraGrid1.DataSource = ultraDataSource1;

			string sSQL = string.Format("Exec GeneraMayorGeneral '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtInicial.Text.ToString(), this.txtFinal.Text.ToString(), bRango, this.txtCuenta.Text.ToString(), bCta);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQL);

			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);			
			this.ultraGrid1.Rows.CollapseAll(true);

			string sSQLSumatorias = string.Format("Exec SumatoriaMayorGeneral '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtInicial.Text.ToString(), this.txtFinal.Text.ToString(), bRango, this.txtCuenta.Text.ToString(), bCta);
			SqlDataReader drSumatorias = Funcion.rEscalarSQL(cdsSeteoF, sSQLSumatorias, true);
			drSumatorias.Read();

			if (drSumatorias.HasRows)
			{
				this.txtTotalDebe.Value = drSumatorias.GetDecimal(0);
				this.txtTotalHaber.Value = drSumatorias.GetDecimal(1);
				this.txtSaldoTotal.Value = drSumatorias.GetDecimal(2);
			}
			drSumatorias.Close();
	
			this.Cursor = Cursors.Default;

			if (this.ultraGrid1.Rows.Count > 0) this.btnImprimir.Enabled = true; else this.btnImprimir.Enabled = false;
		}

		private void btnGeneraBalance_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime)this.dtBCDesde.Value;
			DateTime dtHasta = (DateTime)this.dtBCHasta.Value;

			if (this.txtBCInicial.Text.ToString().Length > 0 && this.txtBCFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBCFinal.Focus();
				return;
			}

			if (this.txtBCInicial.Text.ToString().Length == 0 && this.txtBCFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBCInicial.Focus();
				return;
			}
		
			bool bRango = false;

			if (this.txtBCInicial.Text.ToString().Length > 0 && this.txtBCFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtBCInicial.Text.ToString();
				string sFinal = this.txtBCFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (int.Parse(sInicial) > int.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBCInicial.Focus();
					return;
				}

				bRango = true;
			}

			string sSQL = string.Format("Exec Balances '{0}', '{1}', 0, '{2}', '{3}', {4}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBCInicial.Text.ToString(), this.txtBCFinal.Text.ToString(), bRango);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQL;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsCuentas.DataSource = oDTDetalle;
			}	
			this.Cursor = Cursors.Default;
		}

		private void btnImprimirBR_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaDesde = (DateTime)this.dtBRDesde.Value;
			DateTime dtFechaHasta = (DateTime)this.dtBRHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Desde";
			discreteVal.Value = dtFechaDesde;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@Hasta";
			discreteVal1.Value = dtFechaHasta;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			Reporte miReporte = new Reporte("BalanceDeResultados.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Balance De Resultados");
			miReporte.Show();
		}

		private void btnImprimirBG_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaDesde = (DateTime)this.dtBGDesde.Value;
			DateTime dtFechaHasta = (DateTime)this.dtBGHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Desde";
			discreteVal.Value = dtFechaDesde;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@Hasta";
			discreteVal1.Value = dtFechaHasta;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			Reporte miReporte = new Reporte("BalanceGeneral.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Balance General");
			miReporte.Show();
		}

		private void btnImprimirBC_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaDesde = (DateTime)this.dtBCDesde.Value;
			DateTime dtFechaHasta = (DateTime)this.dtBCHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Desde";
			discreteVal.Value = dtFechaDesde;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@Hasta";
			discreteVal1.Value = dtFechaHasta;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramFieldInicial = new ParameterField ();
			ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
			paramFieldInicial.ParameterFieldName = "@Inicial";
			discreteValInicial.Value = this.txtBCInicial.Text.ToString(); 
			paramFieldInicial.CurrentValues.Add (discreteValInicial);
			paramFields.Add (paramFieldInicial);

			ParameterField paramFieldFinal = new ParameterField ();
			ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
			paramFieldFinal.ParameterFieldName = "@Final";
			discreteValFinal.Value = this.txtBCFinal.Text.ToString(); 
			paramFieldFinal.CurrentValues.Add (discreteValFinal);
			paramFields.Add (paramFieldFinal);

			Reporte miReporte = new Reporte("BalanceDeComprobacion4.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Balance De Comprobación");
			miReporte.Show();
		}

		private void ultraTabPageControl10_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void dtBCDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtBCHasta.Focus();
		}

		private void dtBCHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnImprimirBC.Focus();
		}

		private void dtBGDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtBGHasta.Focus();
		}

		private void dtBGHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnImprimirBG.Focus();
		}

		private void dtBRDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtBRHasta.Focus();
		}

		private void dtBRHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnImprimirBR.Focus();
		}

		private void btnGeneraBalanceGeneral_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime)this.dtBGDesde.Value;
			DateTime dtHasta = (DateTime)this.dtBGHasta.Value;

			string sSQL = string.Format("Exec Balances '{0}', '{1}', 1, '', '', 0", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQL;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsBalanceGeneral.DataSource = oDTDetalle;
			}	
			this.Cursor = Cursors.Default;
		}

		private void btnGenerarBR_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime)this.dtBRDesde.Value;
			DateTime dtHasta = (DateTime)this.dtBRHasta.Value;

			string sSQL = string.Format("Exec Balances '{0}', '{1}', 2, '', '', 0", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQL;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsBalanceDeResultados.DataSource = oDTDetalle;
			}	
			this.Cursor = Cursors.Default;
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{			
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			//this.utcReportes.Tabs[0].Selected
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.ultraGrid1, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void frmReportesDeContabilidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void txtInicial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBCInicial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBCInicial_TextChanged(object sender, System.EventArgs e)
		{
			if (this.txtBCInicial.Text.Length > 0) 
			{
				this.txtBCFinal.Enabled = true;				
			}
			else if (this.txtBCInicial.Text.Length == 0) 
			{
				this.txtBCFinal.Enabled = false;				
				this.txtBCFinal.Text = "";
			}
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
			
			this.lblEstado.Text = "";

			#region Validacion
			if (this.txtCuentaInicial.Text.ToString().Length > 0 && this.txtCuentaFinal.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuentaFinal.Focus();
				return;
			}

			if (this.txtCuentaInicial.Text.ToString().Length == 0 && this.txtCuentaFinal.Text.ToString().Length > 0)
			{
				MessageBox.Show("Ingrese el Rango Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuentaInicial.Focus();
				return;
			}		

			bool bRango = false;

			if (this.txtCuentaInicial.Text.ToString().Length > 0 && this.txtCuentaFinal.Text.ToString().Length > 0)
			{
				string sInicial = this.txtCuentaInicial.Text.ToString();
				string sFinal = this.txtCuentaFinal.Text.Substring(0, sInicial.Length).ToString();

				sInicial = sInicial.Replace(".", "");
				sFinal = sFinal.Replace(".", "");

				if (Int64.Parse(sInicial) > Int64.Parse(sFinal))
				{
					MessageBox.Show("El Rango Inicial no puede ser mayor que el Rango Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCuentaInicial.Focus();
					return;
				}

				bRango = true;
			}
			
			bool bCta = false;

			if (this.txtCuentas.Text.ToString().Length > 0) bCta = true;
			#endregion Validacion

			this.Cursor = Cursors.WaitCursor;

			Funcion.TiempoInicia();

			int idProyecto = 0;
			
			idCuenta = 0;
			dSaldo = 0.00m;

			if (this.cmbProyecto.ActiveRow != null) idProyecto = (int)this.cmbProyecto.Value;

			#region Mayor General
			this.ultraGrid1.DataSource = ultraDataSource1;
			
			idCuenta = 0;
			dSaldo = 0.00m;

			string sSQL = string.Format("Exec GeneraMayorGeneral '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}, {7}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentaInicial.Text.ToString(), this.txtCuentaFinal.Text.ToString(), bRango, this.txtCuentas.Text.ToString(), bCta, idProyecto);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQL);

			idCuenta = 0;
			dSaldo = 0.00m;

			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, false);			
			
			#region Saldo Final
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idCuenta"].Value != System.DBNull.Value)
				{
					if (idCuenta == 0 || idCuenta != (int)dr.Cells["idCuenta"].Value)
					{
						idCuenta = (int)dr.Cells["idCuenta"].Value;
						dSaldo = Convert.ToDecimal(dr.Cells["SaldoInicial"].Value);
						
						dr.Cells["SaldoIni"].Value = Convert.ToDecimal(dr.Cells["SaldoInicial"].Value);

						dr.Cells["Saldo"].Value = Convert.ToDecimal(dr.Cells["SaldoIni"].Value) + Convert.ToDecimal(dr.Cells["Debe"].Value) - Convert.ToDecimal(dr.Cells["Haber"].Value);

						dSaldo = Convert.ToDecimal(dr.Cells["SaldoIni"].Value) + Convert.ToDecimal(dr.Cells["Debe"].Value) - Convert.ToDecimal(dr.Cells["Haber"].Value);			
					}
					else 
					{				
						dr.Cells["SaldoIni"].Value = dSaldo;

						//if (Convert.ToDecimal(dr.Cells["SaldoIni"].Value) > 0.00m)
						dr.Cells["Saldo"].Value = Convert.ToDecimal(dr.Cells["SaldoIni"].Value) + Convert.ToDecimal(dr.Cells["Debe"].Value) - Convert.ToDecimal(dr.Cells["Haber"].Value);

						dSaldo = Convert.ToDecimal(dr.Cells["SaldoIni"].Value) + Convert.ToDecimal(dr.Cells["Debe"].Value) - Convert.ToDecimal(dr.Cells["Haber"].Value);			
					}
				}
			}
			#endregion Saldo Final
			
			this.ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", false, true);			
			this.ultraGrid1.Rows.CollapseAll(true);
			
			string sSQLSumatorias = string.Format("Exec SumatoriaMayorGeneral '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentaInicial.Text.ToString(), this.txtCuentaFinal.Text.ToString(), bRango, this.txtCuentas.Text.ToString(), bCta);
			SqlDataReader drSumatorias = Funcion.rEscalarSQL(cdsSeteoF, sSQLSumatorias, true);
			drSumatorias.Read();

			if (drSumatorias.HasRows)
			{
				this.txtTotalDebe.Value = drSumatorias.GetDecimal(0);
				this.txtTotalHaber.Value = drSumatorias.GetDecimal(1);
				this.txtSaldoTotal.Value = drSumatorias.GetDecimal(2);
			}
			drSumatorias.Close();			
			#endregion Mayor General

			#region Balance de Comprobacion
			string sSQLBalancedeComprobacion = string.Format("Exec Balances '{0}', '{1}', 0, '{2}', '{3}', {4}, '{5}', {6}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentaInicial.Text.ToString(), this.txtCuentaFinal.Text.ToString(), bRango, this.txtCuentas.Text.ToString(), bCta);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQLBalancedeComprobacion;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsCuentas.DataSource = oDTDetalle;
			}
			#endregion Balance de Comprobacion

			#region Balance General
			string sSQLBalanceGeneral = string.Format("Exec Balances '{0}', '{1}', 1, '{2}', '{3}', {4}, '{5}', {6}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentaInicial.Text.ToString(), this.txtCuentaFinal.Text.ToString(), bRango, this.txtCuentas.Text.ToString(), bCta);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQLBalanceGeneral;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsBalanceGeneral.DataSource = oDTDetalle;

				this.lblEstado.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeBalanceGeneral '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd")));
			}	
			#endregion Balance General

			#region Balance de Resultados
			string sSQLBalancedeResultados = string.Format("Exec Balances '{0}', '{1}', 2, '{2}', '{3}', {4}, '{5}', {6}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentaInicial.Text.ToString(), this.txtCuentaFinal.Text.ToString(), bRango, this.txtCuentas.Text.ToString(), bCta);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQLBalancedeResultados;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.tlsBalanceDeResultados.DataSource = oDTDetalle;
			}	
			#endregion Balance de Resultados

			#region Mayor Condensado
			string sSQLMayorCondensado = string.Format("Exec ResumenDeCuentas '{0}', '{1}', '{2}'", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtCuentas.Text.ToString());

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				oCmdActualiza.CommandText = sSQLMayorCondensado;

				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);

				this.uGridMayorCondensado.DataSource = oDTDetalle;
			}	
			#endregion Mayor Condensado

			this.Cursor = Cursors.Default;

			MessageBox.Show("Mayor Generado " + Funcion.Tiempo("EN").ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			if (this.ultraGrid1.Rows.Count > 0) this.btnImprimir.Enabled = true; else this.btnImprimir.Enabled = false;
		}

		private void txtCuentaInicial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaCuentas Busqueda = new BuscaCuentas(this.txtCuentaInicial.Text.ToString(), 1))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtCuentaInicial.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();						
					}
				}
			}

			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCuentaFinal.Enabled) this.txtCuentaFinal.Focus();
				else if (this.utcReportes.Tabs[0].Selected) this.txtCuentas.Focus();
				else this.btnGenerar.Focus();
			}
		}

		private void txtCuentaFinal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaCuentas Busqueda = new BuscaCuentas(this.txtCuentaFinal.Text.ToString(), 1))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtCuentaFinal.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();						
					}
				}
			}

			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCuentaFinal.Enabled) this.txtCuentaFinal.Focus();
				else if (this.utcReportes.Tabs[0].Selected) this.txtCuentas.Focus();
				else this.btnGenerar.Focus();
			}
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCuentaInicial.Focus();
		}

		private void utcReportes_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
			if (this.utcReportes.Tabs[0].Selected)
			{
				this.txtCuentas.Enabled = true;
				this.btnAsiento.Enabled = true;
				this.btExportar.Enabled = true;
			}
			else 
			{				
				this.btnAsiento.Enabled = false;
				//this.btExportar.Enabled = false;
			}
		}

		private void txtCuentaFinal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCuentaInicial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCuentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaCuentas Busqueda = new BuscaCuentas(this.txtCuentas.Text.ToString(), 1))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtCuentas.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();						
					}
				}
			}

			if (e.KeyCode == Keys.Enter) this.btnGenerar_Click(sender, e);
		}

		private void txtCuentaInicial_TextChanged(object sender, System.EventArgs e)
		{
			if (this.txtCuentaInicial.Text.Length > 0) 
			{
				this.txtCuentaFinal.Enabled = true;
				this.txtCuentas.Enabled = false;
			}
			else if (this.txtCuentaInicial.Text.Length == 0) 
			{
				this.txtCuentaFinal.Enabled = false;
				this.txtCuentas.Enabled = true;
				this.txtCuentaFinal.Text = "";
			}
		}

		private void txtCuentas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		int idCuenta = 0;
		int idCuentaMGC = 0;
		decimal dSaldo = 0.00m;
		decimal dSaldoMGC = 0.00m;

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
//			if (idCuenta == 0 || idCuenta != (int)e.Row.Cells["idCuenta"].Value)
//			{
//				idCuenta = (int)e.Row.Cells["idCuenta"].Value;
//				dSaldo = Convert.ToDecimal(e.Row.Cells["SaldoInicial"].Value);
//				//if (Convert.ToDecimal( e.Row.Cells["SaldoIni"].Value == 0.00) e.Row.Cells["SaldoIni"].Value = 0.00;
//				e.Row.Cells["SaldoIni"].Value = Convert.ToDecimal(e.Row.Cells["SaldoInicial"].Value);//dSaldo;
//
//				e.Row.Cells["Saldo"].Value = Convert.ToDecimal(e.Row.Cells["SaldoIni"].Value) + Convert.ToDecimal(e.Row.Cells["Debe"].Value) - Convert.ToDecimal(e.Row.Cells["Haber"].Value);
//
//				dSaldo = Convert.ToDecimal(e.Row.Cells["SaldoIni"].Value) + Convert.ToDecimal(e.Row.Cells["Debe"].Value) - Convert.ToDecimal(e.Row.Cells["Haber"].Value);			
//			}
//			else 
//			{				
//				e.Row.Cells["SaldoIni"].Value = dSaldo;
//
//				e.Row.Cells["Saldo"].Value = Convert.ToDecimal(e.Row.Cells["SaldoIni"].Value) + Convert.ToDecimal(e.Row.Cells["Debe"].Value) - Convert.ToDecimal(e.Row.Cells["Haber"].Value);
//
//				dSaldo = Convert.ToDecimal(e.Row.Cells["SaldoIni"].Value) + Convert.ToDecimal(e.Row.Cells["Debe"].Value) - Convert.ToDecimal(e.Row.Cells["Haber"].Value);			
//			}
		}

		private void btnVacio_Click(object sender, System.EventArgs e)
		{
			this.cmbProyecto.Value = System.DBNull.Value;
		}

		private void uGridMayorCondensado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridMayorCondensado_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
//			if (idCuentaMGC == 0 || idCuentaMGC != (int)e.Row.Cells["idCuenta"].Value)
//			{
//				idCuentaMGC = (int)e.Row.Cells["idCuenta"].Value;
//				dSaldoMGC = Convert.ToDecimal(e.Row.Cells["SaldoInicial"].Value;
//			}

			//e.Row.Cells["SaldoInicial"].Value = dSaldoMGC;

			e.Row.Cells["Saldo"].Value = Convert.ToDecimal(e.Row.Cells["SaldoInicial"].Value) + Convert.ToDecimal(e.Row.Cells["Debe"].Value) - Convert.ToDecimal(e.Row.Cells["Haber"].Value);

			//dSaldoMGC = dSaldoMGC + Convert.ToDecimal(e.Row.Cells["Debe"].Value - Convert.ToDecimal(e.Row.Cells["Haber"].Value;			
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{

		}
	}
}
