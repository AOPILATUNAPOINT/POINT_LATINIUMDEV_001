using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCobranzaAsignacionCartera.
	/// </summary>
	public class frmCobranzaAsignacionCartera : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAsignacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAsignado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label lblContadorPagos;
		private System.Windows.Forms.Label lblContadorOperaciones;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCuotas;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMHasta;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private System.Windows.Forms.Button btnSeleccionar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSeleccionar;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFPagos;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIPagos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGriPagos;
		private System.Windows.Forms.Button btnVerPagos;
		private System.Windows.Forms.Button btnPagos;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnExportarP;
		private System.Windows.Forms.Button btnEmpresa;
		private System.Windows.Forms.Label label11;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbProvincia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCobranzaAsignacionCartera()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCobranzaAsignacionCartera));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultima_Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasVencidas");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVencido");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InteresVencido");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("APagar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_DiasVencidos");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_PorVencer");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Vencido");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Mora");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Gastos");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Valor");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Cobrado");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("A_Usuario");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Provincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "APagar", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "APagar", 14, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mora", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mora", 12, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Gastos", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Gastos", 13, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "InteresVencido", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "InteresVencido", 11, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalVencido", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalVencido", 10, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_Vencido", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_Vencido", 19, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_Mora", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_Mora", 20, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_Gastos", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_Gastos", 21, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_Valor", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_Valor", 22, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "A_Cobrado", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "A_Cobrado", 23, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasVencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InteresVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("APagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_DiasVencidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_PorVencer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Provincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje_Comision");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 10, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 13, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje_Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo_Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo_Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero_Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero_Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Capital_Cobrado");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes_Mora");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Honorario_Cobrado");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago_Voluntario");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Capital_Condonado");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes_Capital_Condonado");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes_Mora_Condonado");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Honorario_Condonado");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Forma_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Administrativo");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCobrado");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Capital_Cobrado", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Capital_Cobrado", 7, true);
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 8, true);
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes_Mora", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes_Mora", 9, true);
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Honorario_Cobrado", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Honorario_Cobrado", 10, true);
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago_Voluntario", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago_Voluntario", 11, true);
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Capital_Condonado", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Capital_Condonado", 12, true);
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes_Capital_Condonado", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes_Capital_Condonado", 13, true);
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes_Mora_Condonado", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes_Mora_Condonado", 14, true);
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Honorario_Condonado", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Honorario_Condonado", 15, true);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Administrativo", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Administrativo", 18, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalCobrado", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalCobrado", 19, true);
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero_Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero_Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Vencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Capital_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Honorario_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago_Voluntario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Capital_Condonado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes_Capital_Condonado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes_Mora_Condonado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Honorario_Condonado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Forma_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalCobrado");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobranzaExterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobranzaExterna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasVencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InteresVencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_DiasVencidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_PorVencer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("A_Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Provincia");
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.txtSeleccionar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContadorOperaciones = new System.Windows.Forms.Label();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.optAsignado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAsignacion = new System.Windows.Forms.Button();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblContadorPagos = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnExportarP = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnPagos = new System.Windows.Forms.Button();
			this.btnVerPagos = new System.Windows.Forms.Button();
			this.uGriPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFPagos = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtIPagos = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.label1 = new System.Windows.Forms.Label();
			this.optCuotas = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtDMDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDMHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnEmpresa = new System.Windows.Forms.Button();
			this.optRegion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label11 = new System.Windows.Forms.Label();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.TbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtSeleccionar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignado)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGriPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			this.SuspendLayout();
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.chkLocal);
			this.TbBasico.Controls.Add(this.btnSeleccionar);
			this.TbBasico.Controls.Add(this.txtSeleccionar);
			this.TbBasico.Controls.Add(this.lblContadorOperaciones);
			this.TbBasico.Controls.Add(this.uGridDocumentos);
			this.TbBasico.Controls.Add(this.cmbCargos);
			this.TbBasico.Controls.Add(this.optAsignado);
			this.TbBasico.Controls.Add(this.btnAsignacion);
			this.TbBasico.Location = new System.Drawing.Point(1, 23);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(1220, 406);
			// 
			// chkLocal
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			appearance1.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocal.Appearance = appearance1;
			this.chkLocal.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocal.Checked = true;
			this.chkLocal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocal.Location = new System.Drawing.Point(40, 40);
			this.chkLocal.Name = "chkLocal";
			this.chkLocal.Size = new System.Drawing.Size(13, 16);
			this.chkLocal.TabIndex = 987;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSeleccionar.Location = new System.Drawing.Point(296, 8);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(80, 24);
			this.btnSeleccionar.TabIndex = 773;
			this.btnSeleccionar.Text = "&Seleccionar";
			this.btnSeleccionar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// txtSeleccionar
			// 
			this.txtSeleccionar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSeleccionar.Location = new System.Drawing.Point(208, 9);
			this.txtSeleccionar.MaxValue = 3600000;
			this.txtSeleccionar.MinValue = 0;
			this.txtSeleccionar.Name = "txtSeleccionar";
			this.txtSeleccionar.PromptChar = ' ';
			this.txtSeleccionar.Size = new System.Drawing.Size(72, 22);
			this.txtSeleccionar.TabIndex = 772;
			// 
			// lblContadorOperaciones
			// 
			this.lblContadorOperaciones.AutoSize = true;
			this.lblContadorOperaciones.BackColor = System.Drawing.Color.Transparent;
			this.lblContadorOperaciones.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorOperaciones.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorOperaciones.Location = new System.Drawing.Point(400, 10);
			this.lblContadorOperaciones.Name = "lblContadorOperaciones";
			this.lblContadorOperaciones.Size = new System.Drawing.Size(0, 20);
			this.lblContadorOperaciones.TabIndex = 771;
			this.lblContadorOperaciones.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance3;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 40;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Ult. Gestion";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Header.VisiblePosition = 10;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 11;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Header.VisiblePosition = 12;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 13;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Header.Caption = "Dias Mora";
			ultraGridColumn9.Header.VisiblePosition = 14;
			ultraGridColumn9.Width = 70;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Header.Caption = "Cuotas Venc";
			ultraGridColumn10.Header.VisiblePosition = 15;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Cap Vencido";
			ultraGridColumn11.Header.VisiblePosition = 16;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance9;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Int Vencido";
			ultraGridColumn12.Header.VisiblePosition = 17;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance10;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 18;
			ultraGridColumn13.Width = 90;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance11;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 20;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.Header.VisiblePosition = 19;
			ultraGridColumn16.Header.Caption = "A";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 30;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Asig Fecha";
			ultraGridColumn17.Header.VisiblePosition = 21;
			ultraGridColumn17.Width = 110;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance12;
			ultraGridColumn18.Header.Caption = "Asig Dias Mora";
			ultraGridColumn18.Header.VisiblePosition = 22;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance13;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 24;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance14;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Asig Vencido";
			ultraGridColumn20.Header.VisiblePosition = 27;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance15;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Asig Mora";
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Width = 100;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance16;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Asig Gastos";
			ultraGridColumn22.Header.VisiblePosition = 25;
			ultraGridColumn22.Width = 100;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance17;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Asig Valor";
			ultraGridColumn23.Header.VisiblePosition = 26;
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance18;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Cobrado";
			ultraGridColumn24.Header.VisiblePosition = 28;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance19;
			ultraGridColumn25.Header.Caption = "Asig Usuario";
			ultraGridColumn25.Header.VisiblePosition = 29;
			ultraGridColumn25.Width = 110;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 31;
			ultraGridColumn27.Width = 200;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 30;
			ultraGridColumn28.Width = 150;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "Ultima Gestión";
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Width = 150;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Gestión";
			ultraGridColumn31.Header.VisiblePosition = 8;
			ultraGridColumn31.Width = 150;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "Estado";
			ultraGridColumn32.Header.VisiblePosition = 9;
			ultraGridColumn32.Width = 150;
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
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance20;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance21;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance22;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance23;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance24;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance25;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance26;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance27;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance28;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance29;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings10});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance30;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance33.FontData.BoldAsString = "True";
			appearance33.FontData.Name = "Arial";
			appearance33.FontData.SizeInPoints = 8F;
			appearance33.ForeColor = System.Drawing.Color.White;
			appearance33.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance33;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance34.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance34.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 40);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1208, 365);
			this.uGridDocumentos.TabIndex = 764;
			this.uGridDocumentos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(752, 448);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(24, 8);
			this.cmbCargos.TabIndex = 763;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// optAsignado
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAsignado.Appearance = appearance37;
			this.optAsignado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAsignado.CheckedIndex = 0;
			this.optAsignado.ItemAppearance = appearance38;
			this.optAsignado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "No Asignado";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Asignado";
			this.optAsignado.Items.Add(valueListItem1);
			this.optAsignado.Items.Add(valueListItem2);
			this.optAsignado.ItemSpacingVertical = 10;
			this.optAsignado.Location = new System.Drawing.Point(8, 8);
			this.optAsignado.Name = "optAsignado";
			this.optAsignado.Size = new System.Drawing.Size(184, 24);
			this.optAsignado.TabIndex = 678;
			this.optAsignado.Text = "No Asignado";
			this.optAsignado.ValueChanged += new System.EventHandler(this.optAsignado_ValueChanged);
			// 
			// btnAsignacion
			// 
			this.btnAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsignacion.Enabled = false;
			this.btnAsignacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignacion.Location = new System.Drawing.Point(1128, 8);
			this.btnAsignacion.Name = "btnAsignacion";
			this.btnAsignacion.Size = new System.Drawing.Size(80, 24);
			this.btnAsignacion.TabIndex = 679;
			this.btnAsignacion.Text = "&Asignar";
			this.btnAsignacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignacion.Click += new System.EventHandler(this.btnAsignacion_Click);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.lblContadorPagos);
			this.ultraTabPageControl10.Controls.Add(this.label4);
			this.ultraTabPageControl10.Controls.Add(this.label5);
			this.ultraTabPageControl10.Controls.Add(this.dtHasta);
			this.ultraTabPageControl10.Controls.Add(this.dtDesde);
			this.ultraTabPageControl10.Controls.Add(this.uGridPagos);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(1220, 406);
			// 
			// lblContadorPagos
			// 
			this.lblContadorPagos.AutoSize = true;
			this.lblContadorPagos.BackColor = System.Drawing.Color.Transparent;
			this.lblContadorPagos.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorPagos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorPagos.Location = new System.Drawing.Point(424, 8);
			this.lblContadorPagos.Name = "lblContadorPagos";
			this.lblContadorPagos.Size = new System.Drawing.Size(0, 20);
			this.lblContadorPagos.TabIndex = 770;
			this.lblContadorPagos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(176, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 769;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 17);
			this.label5.TabIndex = 768;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance39;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 767;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance40;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 766;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// uGridPagos
			// 
			this.uGridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagos.DataSource = this.ultraDataSource2;
			appearance41.BackColor = System.Drawing.Color.White;
			this.uGridPagos.DisplayLayout.Appearance = appearance41;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn36.CellAppearance = appearance42;
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Width = 120;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Width = 70;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Width = 140;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn39.CellAppearance = appearance43;
			ultraGridColumn39.Header.VisiblePosition = 5;
			ultraGridColumn39.Width = 90;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn40.Header.Caption = "Forma De Pago";
			ultraGridColumn40.Header.VisiblePosition = 6;
			ultraGridColumn40.Width = 120;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn41.CellAppearance = appearance44;
			ultraGridColumn41.Header.VisiblePosition = 7;
			ultraGridColumn41.Width = 90;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn42.Header.VisiblePosition = 8;
			ultraGridColumn42.Width = 250;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn44.CellAppearance = appearance45;
			ultraGridColumn44.Header.VisiblePosition = 9;
			ultraGridColumn44.Width = 90;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance46;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 10;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn46.CellAppearance = appearance47;
			ultraGridColumn46.Header.Caption = "Dias Mora";
			ultraGridColumn46.Header.VisiblePosition = 11;
			ultraGridColumn46.Width = 75;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance48;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "% Comisión";
			ultraGridColumn47.Header.VisiblePosition = 12;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance49;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "Comisión";
			ultraGridColumn48.Header.VisiblePosition = 13;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn49.Header.VisiblePosition = 14;
			ultraGridColumn49.Width = 120;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn49});
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance50;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance51;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings11,
																																															summarySettings12});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridPagos.DisplayLayout.GroupByBox.Hidden = true;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagos.DisplayLayout.Override.ActiveRowAppearance = appearance52;
			this.uGridPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagos.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagos.DisplayLayout.Override.CellAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance55.FontData.BoldAsString = "True";
			appearance55.FontData.Name = "Arial";
			appearance55.FontData.SizeInPoints = 8F;
			appearance55.ForeColor = System.Drawing.Color.White;
			appearance55.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagos.DisplayLayout.Override.HeaderAppearance = appearance55;
			this.uGridPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance56.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance56.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowAlternateAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance58;
			this.uGridPagos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridPagos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagos.Location = new System.Drawing.Point(6, 40);
			this.uGridPagos.Name = "uGridPagos";
			this.uGridPagos.Size = new System.Drawing.Size(1208, 360);
			this.uGridPagos.TabIndex = 765;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn47});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btnExportarP);
			this.ultraTabPageControl1.Controls.Add(this.lblContador);
			this.ultraTabPageControl1.Controls.Add(this.btnPagos);
			this.ultraTabPageControl1.Controls.Add(this.btnVerPagos);
			this.ultraTabPageControl1.Controls.Add(this.uGriPagos);
			this.ultraTabPageControl1.Controls.Add(this.label9);
			this.ultraTabPageControl1.Controls.Add(this.label10);
			this.ultraTabPageControl1.Controls.Add(this.dtFPagos);
			this.ultraTabPageControl1.Controls.Add(this.dtIPagos);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1220, 406);
			// 
			// btnExportarP
			// 
			this.btnExportarP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportarP.Image = ((System.Drawing.Image)(resources.GetObject("btnExportarP.Image")));
			this.btnExportarP.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarP.Location = new System.Drawing.Point(1024, 16);
			this.btnExportarP.Name = "btnExportarP";
			this.btnExportarP.Size = new System.Drawing.Size(80, 24);
			this.btnExportarP.TabIndex = 1004;
			this.btnExportarP.Text = "&Excel";
			this.btnExportarP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportarP.Click += new System.EventHandler(this.btnExportarP_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(464, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 1003;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPagos
			// 
			this.btnPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnPagos.Image")));
			this.btnPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPagos.Location = new System.Drawing.Point(1112, 16);
			this.btnPagos.Name = "btnPagos";
			this.btnPagos.Size = new System.Drawing.Size(96, 24);
			this.btnPagos.TabIndex = 1002;
			this.btnPagos.Text = "Expotar Pagos";
			this.btnPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnPagos.Click += new System.EventHandler(this.btnPagos_Click);
			// 
			// btnVerPagos
			// 
			this.btnVerPagos.Image = ((System.Drawing.Image)(resources.GetObject("btnVerPagos.Image")));
			this.btnVerPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerPagos.Location = new System.Drawing.Point(368, 16);
			this.btnVerPagos.Name = "btnVerPagos";
			this.btnVerPagos.Size = new System.Drawing.Size(80, 24);
			this.btnVerPagos.TabIndex = 775;
			this.btnVerPagos.Text = "Ver Pagos";
			this.btnVerPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVerPagos.Click += new System.EventHandler(this.btnVerPagos_Click);
			// 
			// uGriPagos
			// 
			this.uGriPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGriPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGriPagos.DataSource = this.ultraDataSource4;
			appearance59.BackColor = System.Drawing.Color.White;
			this.uGriPagos.DisplayLayout.Appearance = appearance59;
			ultraGridColumn50.Header.Caption = "Tipo";
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn51.Header.Caption = "Tipo Credito";
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridColumn52.Header.Caption = "Cedula";
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridColumn52.Width = 110;
			ultraGridColumn53.Header.Caption = "# Operación";
			ultraGridColumn53.Header.VisiblePosition = 3;
			ultraGridColumn54.Header.Caption = "Fecha Vencimiento";
			ultraGridColumn54.Header.VisiblePosition = 4;
			ultraGridColumn54.Width = 120;
			ultraGridColumn55.Header.Caption = "Fecha Pago";
			ultraGridColumn55.Header.VisiblePosition = 5;
			ultraGridColumn56.Header.Caption = "#Cuota";
			ultraGridColumn56.Header.VisiblePosition = 6;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance60;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.Caption = "Valor Capital Cobrado";
			ultraGridColumn57.Header.VisiblePosition = 8;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance61;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 9;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance62;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 10;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance63;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.Caption = "Valor Honorario Cobrado";
			ultraGridColumn60.Header.VisiblePosition = 11;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance64;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.Caption = "Pago Voluntario";
			ultraGridColumn61.Header.VisiblePosition = 12;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance65;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "Valor Capital Condonado";
			ultraGridColumn62.Header.VisiblePosition = 13;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance66;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "Interes Capital Condonado";
			ultraGridColumn63.Header.VisiblePosition = 14;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance67;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.Caption = "Interes Mora Condonado";
			ultraGridColumn64.Header.VisiblePosition = 15;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance68;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "Valor Honorario Condonado";
			ultraGridColumn65.Header.VisiblePosition = 16;
			ultraGridColumn66.Header.VisiblePosition = 17;
			ultraGridColumn67.Header.Caption = "Tipo Pago";
			ultraGridColumn67.Header.VisiblePosition = 18;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance69;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.Caption = "Valor Administrativo";
			ultraGridColumn68.Header.VisiblePosition = 19;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance70;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "Total Cobrado";
			ultraGridColumn69.Header.VisiblePosition = 7;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69});
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance71;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings13,
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGriPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGriPagos.DisplayLayout.GroupByBox.Hidden = true;
			appearance82.ForeColor = System.Drawing.Color.Black;
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGriPagos.DisplayLayout.Override.ActiveRowAppearance = appearance82;
			this.uGriPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGriPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance83.BackColor = System.Drawing.Color.Transparent;
			this.uGriPagos.DisplayLayout.Override.CardAreaAppearance = appearance83;
			appearance84.ForeColor = System.Drawing.Color.Black;
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGriPagos.DisplayLayout.Override.CellAppearance = appearance84;
			appearance85.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance85.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance85.FontData.BoldAsString = "True";
			appearance85.FontData.Name = "Arial";
			appearance85.FontData.SizeInPoints = 8F;
			appearance85.ForeColor = System.Drawing.Color.White;
			appearance85.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGriPagos.DisplayLayout.Override.HeaderAppearance = appearance85;
			this.uGriPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance86.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance86.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriPagos.DisplayLayout.Override.RowAlternateAppearance = appearance86;
			appearance87.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance87.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance87.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriPagos.DisplayLayout.Override.RowSelectorAppearance = appearance87;
			appearance88.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance88.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriPagos.DisplayLayout.Override.SelectedRowAppearance = appearance88;
			this.uGriPagos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGriPagos.Enabled = false;
			this.uGriPagos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGriPagos.Location = new System.Drawing.Point(8, 48);
			this.uGriPagos.Name = "uGriPagos";
			this.uGriPagos.Size = new System.Drawing.Size(1208, 352);
			this.uGriPagos.TabIndex = 774;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn52.DataType = typeof(System.DateTime);
			ultraDataColumn53.DataType = typeof(System.DateTime);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(189, 20);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(33, 17);
			this.label9.TabIndex = 773;
			this.label9.Text = "Hasta";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(16, 20);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(35, 17);
			this.label10.TabIndex = 772;
			this.label10.Text = "Desde";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFPagos
			// 
			appearance89.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFPagos.Appearance = appearance89;
			this.dtFPagos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton3.Caption = "Today";
			dateButton3.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFPagos.DateButtons.Add(dateButton3);
			this.dtFPagos.Format = "dd/MM/yyyy";
			this.dtFPagos.Location = new System.Drawing.Point(239, 18);
			this.dtFPagos.Name = "dtFPagos";
			this.dtFPagos.NonAutoSizeHeight = 23;
			this.dtFPagos.Size = new System.Drawing.Size(104, 21);
			this.dtFPagos.SpinButtonsVisible = true;
			this.dtFPagos.TabIndex = 771;
			this.dtFPagos.Value = ((object)(resources.GetObject("dtFPagos.Value")));
			// 
			// dtIPagos
			// 
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtIPagos.Appearance = appearance90;
			this.dtIPagos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton4.Caption = "Today";
			dateButton4.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtIPagos.DateButtons.Add(dateButton4);
			this.dtIPagos.Format = "dd/MM/yyyy";
			this.dtIPagos.Location = new System.Drawing.Point(68, 18);
			this.dtIPagos.Name = "dtIPagos";
			this.dtIPagos.NonAutoSizeHeight = 23;
			this.dtIPagos.Size = new System.Drawing.Size(104, 21);
			this.dtIPagos.SpinButtonsVisible = true;
			this.dtIPagos.TabIndex = 770;
			this.dtIPagos.Value = ((object)(resources.GetObject("dtIPagos.Value")));
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(256, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 17);
			this.label3.TabIndex = 677;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// txtBuscar
			// 
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance91;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(312, 80);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 676;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1136, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 675;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(536, 80);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 24);
			this.btnVer.TabIndex = 674;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance92.FontData.Name = "Tahoma";
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance92;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Controls.Add(this.ultraTabPageControl1);
			this.utcReportes.Location = new System.Drawing.Point(8, 112);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(1224, 432);
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.Flat;
			this.utcReportes.TabIndex = 680;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Asignación De Cartera";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Lista De Pagos";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Archivo Pagos";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1220, 406);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(520, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(83, 17);
			this.label1.TabIndex = 996;
			this.label1.Text = "Cuotas Pagadas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optCuotas
			// 
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.optCuotas.Appearance = appearance93;
			this.optCuotas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optCuotas.CheckedIndex = 0;
			this.optCuotas.ItemAppearance = appearance94;
			this.optCuotas.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todos";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "0-1";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "2 - 6";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "7 - 15";
			valueListItem7.DataValue = 4;
			valueListItem7.DisplayText = ">= 16";
			this.optCuotas.Items.Add(valueListItem3);
			this.optCuotas.Items.Add(valueListItem4);
			this.optCuotas.Items.Add(valueListItem5);
			this.optCuotas.Items.Add(valueListItem6);
			this.optCuotas.Items.Add(valueListItem7);
			this.optCuotas.ItemSpacingVertical = 10;
			this.optCuotas.Location = new System.Drawing.Point(608, 3);
			this.optCuotas.Name = "optCuotas";
			this.optCuotas.Size = new System.Drawing.Size(320, 26);
			this.optCuotas.TabIndex = 995;
			this.optCuotas.Text = "Todos";
			this.optCuotas.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(264, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 17);
			this.label8.TabIndex = 998;
			this.label8.Text = "Banco";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEntidadFinanciera
			// 
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance95;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.Width = 208;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn70,
																										 ultraGridColumn71});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(312, 6);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(200, 21);
			this.cmbEntidadFinanciera.TabIndex = 997;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			this.cmbEntidadFinanciera.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEntidadFinanciera_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtDMDesde);
			this.groupBox1.Controls.Add(this.txtDMHasta);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(248, 64);
			this.groupBox1.TabIndex = 999;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Rango";
			// 
			// txtDMDesde
			// 
			this.txtDMDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMDesde.Location = new System.Drawing.Point(48, 24);
			this.txtDMDesde.MaxValue = 3600000;
			this.txtDMDesde.MinValue = 1;
			this.txtDMDesde.Name = "txtDMDesde";
			this.txtDMDesde.PromptChar = ' ';
			this.txtDMDesde.Size = new System.Drawing.Size(72, 22);
			this.txtDMDesde.TabIndex = 671;
			this.txtDMDesde.Value = 1;
			// 
			// txtDMHasta
			// 
			this.txtDMHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMHasta.Location = new System.Drawing.Point(168, 24);
			this.txtDMHasta.MaxValue = 3600000;
			this.txtDMHasta.MinValue = 1;
			this.txtDMHasta.Name = "txtDMHasta";
			this.txtDMHasta.PromptChar = ' ';
			this.txtDMHasta.Size = new System.Drawing.Size(72, 22);
			this.txtDMHasta.TabIndex = 673;
			this.txtDMHasta.Value = 1;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(128, 26);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(33, 17);
			this.label6.TabIndex = 674;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 26);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(35, 17);
			this.label7.TabIndex = 672;
			this.label7.Text = "Desde";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1136, 40);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(96, 24);
			this.btnGenerar.TabIndex = 1001;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(792, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(47, 17);
			this.label2.TabIndex = 1003;
			this.label2.Text = "Empresa";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEmpresa
			// 
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresa.Appearance = appearance96;
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.ultraDataSource3;
			ultraGridColumn72.Header.VisiblePosition = 0;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 1;
			ultraGridColumn73.Width = 198;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn72,
																										 ultraGridColumn73});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(848, 40);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(200, 21);
			this.cmbEmpresa.TabIndex = 1005;
			this.cmbEmpresa.ValueMember = "idCobranzaExterna";
			this.cmbEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEmpresa_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn68.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn68,
																																 ultraDataColumn69});
			// 
			// btnEmpresa
			// 
			this.btnEmpresa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnEmpresa.Image = ((System.Drawing.Image)(resources.GetObject("btnEmpresa.Image")));
			this.btnEmpresa.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEmpresa.Location = new System.Drawing.Point(1056, 38);
			this.btnEmpresa.Name = "btnEmpresa";
			this.btnEmpresa.Size = new System.Drawing.Size(25, 24);
			this.btnEmpresa.TabIndex = 1006;
			this.btnEmpresa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEmpresa.Visible = false;
			this.btnEmpresa.Click += new System.EventHandler(this.btnEmpresa_Click);
			// 
			// optRegion
			// 
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegion.Appearance = appearance97;
			this.optRegion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegion.CheckedIndex = 0;
			this.optRegion.ItemAppearance = appearance98;
			this.optRegion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Todos";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Sierra";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Costa";
			this.optRegion.Items.Add(valueListItem8);
			this.optRegion.Items.Add(valueListItem9);
			this.optRegion.Items.Add(valueListItem10);
			this.optRegion.ItemSpacingVertical = 10;
			this.optRegion.Location = new System.Drawing.Point(312, 40);
			this.optRegion.Name = "optRegion";
			this.optRegion.Size = new System.Drawing.Size(200, 26);
			this.optRegion.TabIndex = 1007;
			this.optRegion.Text = "Todos";
			this.optRegion.ValueChanged += new System.EventHandler(this.ultraOptionSet1_ValueChanged);
			// 
			// cmbProvincia
			// 
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProvincia.Appearance = appearance99;
			this.cmbProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvincia.DataSource = this.ultraDataSource5;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridColumn75.Width = 200;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75});
			this.cmbProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbProvincia.DisplayMember = "Nombre";
			this.cmbProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvincia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvincia.Location = new System.Drawing.Point(576, 40);
			this.cmbProvincia.Name = "cmbProvincia";
			this.cmbProvincia.Size = new System.Drawing.Size(200, 21);
			this.cmbProvincia.TabIndex = 1009;
			this.cmbProvincia.ValueMember = "idProvincia";
			this.cmbProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProvincia_KeyDown);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn70.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn70,
																																 ultraDataColumn71});
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(520, 42);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(49, 17);
			this.label11.TabIndex = 1008;
			this.label11.Text = "Provincia";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn76.DataType = typeof(System.DateTime);
			ultraDataColumn81.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(int);
			ultraDataColumn85.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(System.Decimal);
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn88.DataType = typeof(System.Decimal);
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn90.DataType = typeof(System.Decimal);
			ultraDataColumn91.DataType = typeof(bool);
			ultraDataColumn92.DataType = typeof(System.DateTime);
			ultraDataColumn93.DataType = typeof(int);
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(System.Decimal);
			ultraDataColumn98.DataType = typeof(System.Decimal);
			ultraDataColumn99.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89,
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99,
																																 ultraDataColumn100,
																																 ultraDataColumn101,
																																 ultraDataColumn102});
			// 
			// frmCobranzaAsignacionCartera
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1240, 550);
			this.Controls.Add(this.cmbProvincia);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.optRegion);
			this.Controls.Add(this.btnEmpresa);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.optCuotas);
			this.Controls.Add(this.utcReportes);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnExcel);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCobranzaAsignacionCartera";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asignación De Cartera";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCobranzaAsignacionCartera_KeyDown);
			this.Load += new System.EventHandler(this.frmCobranzaAsignacionCartera_Load);
			this.TbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtSeleccionar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignado)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGriPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmCobranzaAsignacionCartera_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083203");

			if (!Funcion.Permiso("969", cdsSeteoF))
			{				
				MessageBox.Show("No tiene el acceso  Asignación de Cartera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			if (miAcceso.Editar) 	this.btnAsignacion.Enabled = true;
			if (miAcceso.Nuevo) this.btnEmpresa.Visible = true;
			if (miAcceso.BGenerar) this.btnGenerar.Enabled = true;

			DateTime dtFecha = DateTime.Parse("06/11/2018");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtHasta.Value = DateTime.Today;
			
			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera not in (6,2) Order By idEntidadFinanciera");
			this.cmbEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " select idCobranzaExterna, Empresa from CobranzaExterna");
			this.cmbProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " SELECT idProvincia, Nombre FROM Cre_Provincia ");
			
			this.dtIPagos.CalendarInfo.MinDate = DateTime.Parse("11/07/2019");
			this.dtFPagos.CalendarInfo.MinDate = DateTime.Parse("11/07/2019");
			this.dtIPagos.Value = DateTime.Today;
			this.dtFPagos.Value = DateTime.Today;

		this.cmbEntidadFinanciera.Focus();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.utcReportes.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
			if (this.utcReportes.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridPagos);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Validacion

			if ( (int)this.txtDMDesde.Value > (int) this.txtDMHasta.Value)
			{
				MessageBox.Show("Los dias de mora Inicial no pueden ser mayor al Final  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDMDesde.Focus ();
				return;
			}
			int idEntidadFinanciera  = 0;

			if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;
			int iEmpresa = 0;

			if (this.cmbEmpresa.ActiveRow != null) iEmpresa = (int)this.cmbEmpresa.Value;

			int iProvincia= 0;

			if (this.cmbProvincia.ActiveRow != null) iProvincia = (int)this.cmbProvincia.Value;

			int iCuotaD = 0;
			int iCuotaH = 0;

			if((int)this.optCuotas.Value == 0) 
			{
				iCuotaD = 0; 
				iCuotaH = 50;
			}
			if((int)this.optCuotas.Value == 1) 
			{
				iCuotaD = 0; 
				iCuotaH = 1;
			}
			if((int)this.optCuotas.Value == 2)
			{
				iCuotaD = 2; 
				iCuotaH = 6;
			}
			if((int)this.optCuotas.Value == 3)
			{
				iCuotaD = 7; 
				iCuotaH = 15;
			}
			if((int)this.optCuotas.Value == 4) 
			{
				iCuotaD = 16; 
				iCuotaH = 50;
			}
     #endregion Validacion

			if (this.utcReportes.Tabs[0].Selected)
			{
				string sSQL = string.Format("Exec CobranzasAsignacionCarteraIP {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}", 
					(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
					iCuotaD, iCuotaH, iEmpresa,(int) this.optRegion.Value, iProvincia  );
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);

				this.lblContadorOperaciones.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";
			}

			if (this.utcReportes.Tabs[1].Selected)
			{
				string sSQL = string.Format("Exec CobranzasListaDePagosCarteraIP '{0}', '{1}', '{2}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString());
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPagos);

				this.lblContadorPagos.Text = this.uGridPagos.Rows.Count + " REGISTROS ENCONTRADOS";
			}

			if ((bool) this.optAsignado.Value == true && (int)this.uGridDocumentos.Rows.Count > 0)
			{
				if (miAcceso.Nuevo) this.btnGenerar.Enabled = true; 
			}
			else
			{
				this.btnGenerar.Enabled = false;
			}
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/07/2017");
				this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

				if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
				{
					this.dtHasta.Value = DateTime.Today;
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
				}
				else
				{
					this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
					this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				}

				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);						
			}
			else this.dtHasta.Value = System.DBNull.Value;
		}

		private void btnAsignacion_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbEmpresa, "Seleccione la Empresa a Asignar")) return;
      
			int iCont = 0;
			int iCantidad = 0 ;

			iCantidad = (int) this.uGridDocumentos.Rows.Count ;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if((bool) dr.Cells["Asignado"].Value == true)
				{
					dr.Cells["Asignado"].Value = true;
					iCont = iCont + 1;
				}
			}

			if(iCont == 0)
			{
				MessageBox.Show("Seleccione al menos un Cliente ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
     
			if (DialogResult.Yes == MessageBox.Show( string.Format("¿Esta seguro de Asignar, {0} registros a {1} ", iCont,this.cmbEmpresa.Text), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{						
								
			using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
			{
				oConCobranzas.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
				oCmdActualiza.Connection = oConCobranzas;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{							
					oTransaction = oConCobranzas.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
					{
						if((bool) dr.Cells["Asignado"].Value == true)
						{
							string sSQL = string.Format("Exec CobranzasActualizaAsignacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7},{8}", 
								(int)dr.Cells["idCompra"].Value, 0, (int)dr.Cells["DiasMora"].Value, 
								(Convert.ToDecimal(dr.Cells["CapitalVencido"].Value) + Convert.ToDecimal(dr.Cells["InteresVencido"].Value)),
								Convert.ToDecimal(dr.Cells["Mora"].Value), Convert.ToDecimal(dr.Cells["Gastos"].Value), Convert.ToDecimal(dr.Cells["APagar"].Value), 
								1, (int) this.cmbEmpresa.Value);
							oCmdActualiza.CommandText = sSQL;
							oCmdActualiza.ExecuteNonQuery();
						}
					}

					oTransaction.Commit();	

					oConCobranzas.Close();

				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

					oConCobranzas.Close();					
				}
				finally
				{
					oConCobranzas.Close();
				}			
			}
			}
			
			this.btnVer_Click ( sender, e);
		}

    bool bActualiza = false;
		
		private void uGridDocumentos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{				
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Asignado")
			{
				if( (bool) this.optAsignado.Value  )
				{
					bActualiza = true;

					if ((bool) uGridDocumentos.ActiveRow.Cells["Asignado"].Value  == false)
					{					
						if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Quitar la Asignación Externa?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
						{						
							string sSQL = string.Format(" EXEC ActualizaCarteraExterna {0} ", 
								(int)	this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value );
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);		
							this.btnVer_Click (sender, e);
						}
						else
						{						
							e.Cell.Row.Cells["Asignado"].Value = (bool)e.Cell.OriginalValue;
							//return;
						}					
					}
			
					bActualiza = false;
				}
			}
		}

		private void optAsignado_ValueChanged(object sender, System.EventArgs e)
		{
			if (!(bool)this.optAsignado.Value)
			{
				this.btnSeleccionar.Enabled = true;
				if (miAcceso.Editar) 	this.btnAsignacion.Enabled = true;
				this.btnAsignacion.Enabled = true;
				this.chkLocal.Enabled = true;
			}
			else
			{
				this.btnSeleccionar.Enabled = false;
				this.chkLocal.Enabled = false;
				this.optAsignado.Enabled = true;

			}
			this.btnVer_Click( sender, e);
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			string sArchivos = "";
			if (!Validacion.vbComboVacio(this.cmbEmpresa, "Seleccione la Empresa a Expotar")) return;

			#region Validacion

			int idEntidadFinanciera  = 0;

			int iEmpresa = 0;

			if (this.cmbEmpresa.ActiveRow != null) iEmpresa = (int)this.cmbEmpresa.Value;

			int iProvincia= 0;

			if (this.cmbProvincia.ActiveRow != null) iProvincia = (int)this.cmbProvincia.Value;


			if ( (int)this.txtDMDesde.Value > (int) this.txtDMHasta.Value)
			{
				MessageBox.Show("Los dias de mora Inicial no pueden ser mayor al Final  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDMDesde.Focus ();
				return;
			}


			if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;

			int iCuotaD = 0;
			int iCuotaH = 0;

			if((int)this.optCuotas.Value == 0) 
			{
				iCuotaD = 0; 
				iCuotaH = 50;
			}
			if((int)this.optCuotas.Value == 1) 
			{
				iCuotaD = 0; 
				iCuotaH = 0;
			}
			if((int)this.optCuotas.Value == 5) 
			{
				iCuotaD = 1; 
				iCuotaH = 1;
			}
			if((int)this.optCuotas.Value == 2)
			{
				iCuotaD = 2; 
				iCuotaH = 6;
			}
			if((int)this.optCuotas.Value == 3)
			{
				iCuotaD = 7; 
				iCuotaH = 15;
			}
			if((int)this.optCuotas.Value == 4) 
			{
				iCuotaD = 16; 
				iCuotaH = 50;
			}
			#endregion Validacion

			#region RECOVER
			if( (int) this.cmbEmpresa.Value == 2)
			{
				#region Creditos 

				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = "Recover Creditos.txt";
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Genera Archivo Creditos Recover";
				                                
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverCredito {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Tipo_Credito"].ToString() + "|" 
									+ dr["Numero_Identificacion"].ToString() 
									+ "|" + dr["Numero_Operacion"].ToString() 
									+ "|" + dr["OperacionUnica"].ToString() 
									+ "|" + dr["Fecha_Creacion"].ToString() 
									+ "|" + dr["Fecha_Vencimiento"].ToString() 
									+ "|" + dr["Tasa_Credito"].ToString() 
									+ "|" + dr["Saldo_Operacion"].ToString() 
									+ "|" + dr["Plazo"].ToString()
									+ "|" + dr["Numero_Cuotas_Vencidas"].ToString()
									+ "|" + dr["Numero_Cuota_Mas_Vencida"].ToString()
									+ "|" + dr["Atraso_Promedio"].ToString()
									+ "|" + dr["Atraso_Maximo"].ToString()
									+ "|" + dr["Score_Cobranza"].ToString()
									+ "|" + dr["Segmento"].ToString()
									+ "|" + dr["Clase"].ToString()
									+ "|" + dr["Canal_De_Venta"].ToString()
									+ "|" + dr["Provincia"].ToString()
									+ "|" + dr["Plaza"].ToString()
									+ "|" + dr["Agencia"].ToString()
									+ "|" + dr["Descripcion_Producto"].ToString()
									+ "|" + dr["Tipo_Cliente"].ToString()
									+ "|" + dr["Dias_Mora"].ToString()
									;
							
								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"Creditos/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

					}
					#endregion Pagos
				}
				#endregion Creditos 

				#region Cuotas 

				SaveFileDialog sArchivoCuotas = new SaveFileDialog();

				sArchivoCuotas.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoCuotas.FileName = "Recover Cuotas.txt";
				sArchivoCuotas.OverwritePrompt = true;
				sArchivoCuotas.Title = "Genera Archivo Creditos Recover";
				                                
				if (sArchivoCuotas.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoCuotas.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverCuotas {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Numero_De_Identificacion"].ToString() + "|" 
									+ dr["Numero_Operacion"].ToString() 
									+ "|" + dr["OperacionUnica"].ToString() 
									+ "|" + dr["Fecha_Vencimiento"].ToString() 
									+ "|" + dr["Numero_Cuota"].ToString() 
									+ "|" + dr["Tas_Interes_Mora"].ToString() 
									+ "|" + dr["Capital_Cuota"].ToString() 
									+ "|" + dr["Interes_financiero_Cuota"].ToString() 
									+ "|" + dr["Valor_Mora"].ToString()
									+ "|" + dr["Comision_2_honorarios"].ToString()
									+ "|" + dr["Pago_Voluntario"].ToString()
									+ "|" + dr["Saldo_Capital_Total_Cuota"].ToString()
									+ "|" + dr["Valor_Administrativo"].ToString()
									;
							
								file.WriteLine(iLinea);

								i++;
								//								sArchivos = sArchivos +"Cuotas/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos +"Cuotas/";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

						
					}
					#endregion Cuotas
					
				}

				#endregion Creditos 

				#region Clientes 

				SaveFileDialog sArchivoClientes = new SaveFileDialog();

				sArchivoClientes.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoClientes.FileName = "Recover Clientes.txt";
				sArchivoClientes.OverwritePrompt = true;
				sArchivoClientes.Title = "Genera Archivo Clientes Recover";
				                                
				if (sArchivoClientes.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoClientes.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverClientes {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Numero_De_Identificacion"].ToString() + "|" 
									+ dr["Apellidos"].ToString() 
									+ "|" + dr["Nombres"].ToString() 
									+ "|" + dr["Fecha_Nacimeinto"].ToString() 
									+ "|" + dr["Estado_Civil"].ToString() 
									+ "|" + dr["Genero"].ToString() 
									+ "|" + dr["Actividad_Economica"].ToString() 
									+ "|" + dr["Renta_Mensual"].ToString()
									+ "|" + dr["Tipo_Vivienda"].ToString()
									+ "|" + dr["Empresa_Cliente"].ToString()
									+ "|" + dr["Cargo_Cliente"].ToString()
									+ "|" + dr["Correo_Electronico"].ToString()
									+ "|" + dr["Cargas_Familiares"].ToString() 
									+ "|" + dr["Identificacion_Conyuge"].ToString() 
									+ "|" + dr["Nombre_Conyuge"].ToString() 
									+ "|" + dr["Nombre_Empresa_conyuge"].ToString()
									+ "|" + dr["Cargo_Conyuge"].ToString() ;
							
								file.WriteLine(iLinea);

								i++;
								//								sArchivos = sArchivos +"Clientes/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos +"Clientes/";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

						
					}
					#endregion Clientes
					
				}

				#endregion Creditos 

				#region Referencias 

				SaveFileDialog sArchivoReferencias = new SaveFileDialog();

				sArchivoReferencias.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoReferencias.FileName = "Recover Referencias.txt";
				sArchivoReferencias.OverwritePrompt = true;
				sArchivoReferencias.Title = "Genera Archivo Referencias Recover";
				                                
				if (sArchivoReferencias.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoReferencias.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverReferencias {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Numero_De_Identificacion"].ToString() + "|" 
									+ dr["Nombre_Referencia"].ToString() 
									+ "|" + dr["Telefono_Fijo"].ToString() 
									+ "|" + dr["Telefono_Celular"].ToString() 
									+ "|" + dr["Telefono_Trabajo"].ToString() 
									+ "|" + dr["Extension"].ToString() 
									+ "|" + dr["Tipo_Referencia"].ToString() 
									;
							
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos +"Referencias/";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

						
					}
					#endregion Referencias
					
				}

				#endregion Creditos 

				#region Direcciones 

				SaveFileDialog sArchivoDirecciones = new SaveFileDialog();

				sArchivoDirecciones.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoDirecciones.FileName = "Recover Direcciones.txt";
				sArchivoDirecciones.OverwritePrompt = true;
				sArchivoDirecciones.Title = "Genera Archivo Direcciones Recover";
				                                
				if (sArchivoDirecciones.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoDirecciones.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverDirecciones {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Numero_De_Identificacion"].ToString() + "|" 
									+ dr["Calle_Principal"].ToString() 
									+ "|" + dr["Numero_O_Nomenglatura"].ToString() 
									+ "|" + dr["Numero_Departamento"].ToString() 
									+ "|" + dr["Calle_Secundaria"].ToString() 
									+ "|" + dr["Barrio"].ToString() 
									+ "|" + dr["Canton"].ToString() 
									+ "|" + dr["Parroquia"].ToString() 
									+ "|" + dr["Ciudad"].ToString() 
									+ "|" + dr["Provincia"].ToString() 
									+ "|" + dr["Referencia_Ubicacion"].ToString() 
									+ "|" + dr["Tipo_Direccion"].ToString() 
									+ "|" + dr["Suministro_De_Luz"].ToString() 
									+ "|" + dr["Verificada_Direccion_Domicilio"].ToString() 
									;
							
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos +"Direcciones/";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

					}
					#endregion Pagos
					
				}

				#endregion Direcciones 

				#region Telefono 

				SaveFileDialog sArchivoTelefono = new SaveFileDialog();

				sArchivoTelefono.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoTelefono.FileName = "Recover Telefono.txt";
				sArchivoTelefono.OverwritePrompt = true;
				sArchivoTelefono.Title = "Genera Archivo Telefono Recover";
				                                
				if (sArchivoTelefono.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoTelefono.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverTelefono {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 
									+ dr["Numero_De_Identificacion"].ToString() + "|" 
									+ dr["Numero_De_Telefono"].ToString() 
									+ "|" + dr["Extension"].ToString() 
									+ "|" + dr["Tipo_Telefono"].ToString() 
									+ "|" + dr["Tipo_Origen_Telefono"].ToString()  
									;
							
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos + " Telefono";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

					}
					#endregion Pagos
					
				}

				#endregion Telefono 

				#region Productos 

				SaveFileDialog sArchivoProductos = new SaveFileDialog();

				sArchivoProductos.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoProductos.FileName = "Recover Productos.txt";
				sArchivoProductos.OverwritePrompt = true;
				sArchivoProductos.Title = "Genera Archivo Productos Recover";
				                                
				if (sArchivoProductos.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoProductos.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverPRoductos {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString()
									+ "|" + dr["Numero_Identificacion"].ToString() 
									+ "|" + dr["Nombre"].ToString() 
									+ "|" + dr["Numero_Operacion"].ToString() 
									+ "|" + dr["OperacionUnica"].ToString() 
									+ "|" + dr["Articulo"].ToString() 
									+ "|" + dr["Fecha"].ToString() 
									+ "|" + dr["Total_Credito"].ToString()  
									;
							
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						if (i > 0) sArchivos = sArchivos + " Productos";
						//											MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales

					}
					#endregion Pagos					
				}

				#endregion Productos 

				#region Respaldo cartera Externa 
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;

					try
					{
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						

						string sSQL = string.Format("Exec ConsultaRespaldoCarteraExterna {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}", 
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, iEmpresa,(int) this.optRegion.Value, iProvincia  );
						oCmdActualiza.CommandText = sSQL;
						oCmdActualiza.ExecuteNonQuery();	

						
						oTransaction.Commit();

						MessageBox.Show(" Respaldo De Cartera Externa  Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.lblContadorOperaciones.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";

						MessageBox.Show(string.Format(" Se generaron los siguientes archivos:{0} ", sArchivos), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);


					
					}
					catch(Exception ex)
					{
						try
						{ 
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							MessageBox.Show("El Respaldo De Cartera Externa No Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
						
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							MessageBox.Show("El Respaldo De Cartera Externa No Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
						
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
				#endregion Respaldo cartera Externa 
			}
			#endregion RECOVER

			#region SiCobra
			
			if( (int) this.cmbEmpresa.Value == 3)
			{
				#region SiCobracliente 

				SaveFileDialog SiCobracliente = new SaveFileDialog();

				SiCobracliente.Filter = "Documentos de texto (*.txt)|*.txt";
				SiCobracliente.FileName = "SiCobracliente.txt";
				SiCobracliente.OverwritePrompt = true;
				SiCobracliente.Title = "Genera Archivo SiCobra cliente";
				                                
				if (SiCobracliente.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(SiCobracliente.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraClientes {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 	+	
									dr["Identificacion"].ToString() + "|" 	+	
									dr["Apellidos"].ToString() + "|" 	+	
									dr["Nombres"].ToString() + "|" 	+
									dr["FechaNacimeinto"].ToString() + "|" 	+	
									dr["Genero"].ToString() + "|" 	+	
									dr["EstadoCivil"].ToString() + "|" 	+	
									dr["Nacionalidad"].ToString() + "|" 	+	
									dr["ProfesionOcupacion"].ToString() + "|" 	+	
									dr["NumeroCargas"].ToString() + "|" 	+	
									dr["NombreEmpresa"].ToString() + "|" 	+	
									dr["ActividadEconomica"].ToString() + "|" 	+	
									dr["RUCEmpresa"].ToString() + "|" 	+	
									dr["idProvincia"].ToString() + "|" 	+	
									dr["TipoCliente"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobracliente/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobracliente 

				#region SiCobraConyuge 

				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = "SiCobraConyuge.txt";
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Genera Archivo SiCobra Conyuge";
				                                
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraConyuge {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 
									dr["Tipo_Documento"].ToString() + "|" 	+
									dr["Identificacion"].ToString() + "|" 	+
									dr["TipoIdConyuge"].ToString() + "|" +
									dr["IdentificacionConyuge"].ToString() + "|" +
									dr["ApellidosConyuge"].ToString() + "|"+
									dr["NombresConyuge"].ToString() + "|" 	+
									dr["FechaNacimientoConyuge"].ToString() + "|" 	+
									dr["SexoConyuge"].ToString() + "|" 	+
									dr["EstadoCivilConyuge"].ToString() + "|" 	+
									dr["NacionalidadConyuge"].ToString() + "|" 	+
									dr["Profesion_Ocupacion"].ToString() + "|" +
									dr["ActividadEconomica_conyuge"].ToString() + "|" +	
									dr["NumeroCargas_Conyuge"].ToString() + "|" 	+
									dr["Provincia_Conyuge"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraConyuge/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraConyuge 

				#region SiCobraDireccion

				SaveFileDialog SiCobraDireccion = new SaveFileDialog();

				SiCobraDireccion.Filter = "Documentos de texto (*.txt)|*.txt";
				SiCobraDireccion.FileName = "SiCobraDireccion.txt";
				SiCobraDireccion.OverwritePrompt = true;
				SiCobraDireccion.Title = "Genera Archivo SiCobra Direccion";
				                                
				if (SiCobraDireccion.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(SiCobraDireccion.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraDireccion {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 

									dr["Tipo_Documento"].ToString() + "|" 	+	
									dr["Numero_De_Identificacion"].ToString() + "|" 	+	
									dr["idProvincia"].ToString() + "|" 	+	
									dr["idCanton"].ToString() + "|" 	+	
									dr["idParroquia"].ToString() + "|" 	+	
									dr["idBarrio"].ToString() + "|" 	+	
									dr["Calle_Principal"].ToString() + "|" 	+	
									dr["Numero_O_Nomenglatura"].ToString() + "|" 	+	
									dr["Calle_Secundaria"].ToString() + "|" 	+	
									dr["Referencia_Ubicacion"].ToString() + "|" 	+	
									dr["TIPO_DIRECCION"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraDireccion/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraDireccion 

				#region SiCobraDispositivoUbicacion

				SaveFileDialog SiCobraDispositivoUbicacion = new SaveFileDialog();

				SiCobraDispositivoUbicacion.Filter = "Documentos de texto (*.txt)|*.txt";
				SiCobraDispositivoUbicacion.FileName = "SiCobraDispositivoUbicacion.txt";
				SiCobraDispositivoUbicacion.OverwritePrompt = true;
				SiCobraDispositivoUbicacion.Title = "Genera Archivo  SiCobraDispositivo Ubicacion";
				                                
				if (SiCobraDispositivoUbicacion.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(SiCobraDispositivoUbicacion.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraDispositivoUbicacion {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 
                      
									dr["Tipo_Documento"].ToString() + "|" 	+		
									dr["Numero_De_Identificacion"].ToString() + "|" 	+		
									dr["TipoDireccion"].ToString() + "|" 	+		
									dr["Es_Personal"].ToString() + "|" 	+		
									dr["TipoDispositivo"].ToString() + "|" 	+		
									dr["Valor"].ToString() + "|" 	+		
									dr["Extension"].ToString() + "|" 	+		
									dr["EsUbicable"].ToString() ;

									

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraDispositivoUbicacion/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraDispositivoUbicacion 

				#region SiCobraReferencias

				SaveFileDialog sArchivoSiCobraReferencias = new SaveFileDialog();

				sArchivoSiCobraReferencias.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoSiCobraReferencias.FileName = "SiCobraReferencias.txt";
				sArchivoSiCobraReferencias.OverwritePrompt = true;
				sArchivoSiCobraReferencias.Title = "Genera Archivo SiCobra Referencias";
				                                
				if (sArchivoSiCobraReferencias.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(sArchivoSiCobraReferencias.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraReferencias {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 

									dr["Tipo_Documento"].ToString() + "|" 	+	
									dr["Identificacion"].ToString() + "|" 	+	
									dr["Apellidos"].ToString() + "|" 	+
									dr["Nombre"].ToString() + "|" 	+	
									dr["DireccionDomicilio"].ToString() + "|" 	+	
									dr["TelefonoDomicilio"].ToString() + "|" 	+	
									dr["TelefonoTrabajo"].ToString() + "|" 	+	
									dr["Parentesco"].ToString() + "|" 	+	
									dr["NombreEmpresa"].ToString() + "|" 	+
									dr["DireccionEmpresa"].ToString() + "|" 	+	
									dr["CelularReferencia"].ToString() + "|" 	+	
									dr["EmailReferencia"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraReferencias/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraReferencias 

				#region SiCobraCredito

				SaveFileDialog sArchivoSiCobraCredito = new SaveFileDialog();

				sArchivoSiCobraCredito.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoSiCobraCredito.FileName = "SiCobraCredito.txt";
				sArchivoSiCobraCredito.OverwritePrompt = true;
				sArchivoSiCobraCredito.Title = "Genera Archivo SiCobraCredito";
				                                
				if (sArchivoSiCobraCredito.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(sArchivoSiCobraCredito.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraCredito {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 

									dr["Tipo_Documento"].ToString() + "|" 	+		
									dr["Identificacion"].ToString() + "|" 	+		
									dr["NumeroOperacion"].ToString() + "|" 	+		
									dr["ValorDesembolsado"].ToString() + "|" 	+		
									dr["ValorSaldoCapital"].ToString() + "|" 	+		
									dr["PlazoDeuda"].ToString() + "|" 	+		
									dr["FechaCreacion"].ToString() + "|" 	+		
									dr["FechaDesembolso"].ToString() + "|" 	+		
									dr["TipoCredito"].ToString() + "|" 	+		
									dr["CuotasImpagas"].ToString() + "|" 	+		
									dr["FechaProximoVencimiento"].ToString() + "|" 	+		
									dr["AtrasoMaximo"].ToString() + "|" 	+	
									dr["TipoDocumentoGarante"].ToString() + "|" 	+		
									dr["IdentificacionGarante"].ToString() + "|" 	+	
									dr["IdOficinaCredito"].ToString() + "|" 	+		
									dr["NumeroCuotaImpaga"].ToString() + "|" 	+	
									dr["Atraso_Promedio"].ToString() + "|" 	+		
									dr["NumeroFactura"].ToString() + "|" 	+		
									dr["Producto"].ToString() + "|" 	+		
									dr["ValorCompletoCuota"].ToString() + "|" 	+		
									dr["EstadoCredito"].ToString() + "|" 	+		
									dr["DiasVencido"].ToString() + "|" 	+		
									dr["PagoEnProgreso"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraCredito/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraCredito

				#region SiCobraCuotas

				SaveFileDialog sArchivoSiCobraCuotas = new SaveFileDialog();

				sArchivoSiCobraCuotas.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoSiCobraCuotas.FileName = "sArchivoSiCobraCuotas.txt";
				sArchivoSiCobraCuotas.OverwritePrompt = true;
				sArchivoSiCobraCuotas.Title = "Genera Archivo SiCobraCuotas";
				                                
				if (sArchivoSiCobraCuotas.ShowDialog() == DialogResult.OK)
				{					
					#region carga
					using (StreamWriter file = new StreamWriter(sArchivoSiCobraCuotas.FileName, true))
					{
						#region Normales
						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraCuotas {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}",
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, (int)this.cmbEmpresa.Value)).Rows)
						{
							try
							{
								string iLinea = 
                    
									dr["NumeroOperacion"].ToString() + "|" 	+			
									dr["Numero_Cuota"].ToString() + "|" 	+			
									dr["Interes"].ToString() + "|" 	+			
									dr["Mora"].ToString() + "|" 	+			
									dr["Gestion"].ToString() + "|" 	+			
									dr["Capital"].ToString() + "|" 	+			
									dr["Fecha_Vencimiento"].ToString() + "|" 	+			
									dr["EstadoCuota"].ToString() + "|" 	+			
									dr["ValorReestructurado"].ToString() + "|" 	+			
									dr["ValorExtraCalculoCobranza"].ToString() + "|" 	+			
									dr["ValorExtraNoCalculoCobranza"].ToString() + "|" 	+			
									dr["ValorCuotaInicial"].ToString() + "|" 	+			
									dr["ValorCuota"].ToString() + "|" 	+			
									dr["Dias_Mora_Actual"].ToString() ;

								file.WriteLine(iLinea);

								i++;
			
							}					
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}
					
						if (i > 0) sArchivos = sArchivos +"SiCobraCredito/";
						//													MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						#endregion Normales
					}
					#endregion carga
				}
				#endregion SiCobraCuotas

				#region Respaldo cartera Externa 
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;

					try
					{
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						

						string sSQL = string.Format("Exec ConsultaRespaldoCarteraExterna {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}", 
							(bool)this.optAsignado.Value, this.txtBuscar.Text.ToString(), idEntidadFinanciera, (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value,
							iCuotaD, iCuotaH, iEmpresa,(int) this.optRegion.Value, iProvincia  );
						oCmdActualiza.CommandText = sSQL;
						oCmdActualiza.ExecuteNonQuery();	

						
						oTransaction.Commit();

						MessageBox.Show(" Respaldo De Cartera Externa  Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.lblContadorOperaciones.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";

						MessageBox.Show(string.Format(" Se generaron los siguientes archivos:{0} ", sArchivos), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);


					
					}
					catch(Exception ex)
					{
						try
						{ 
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							MessageBox.Show("El Respaldo De Cartera Externa No Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
						
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							MessageBox.Show("El Respaldo De Cartera Externa No Se Ejecuto Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
						
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
				#endregion Respaldo cartera Externa 

			}
				#endregion SiCobra

		}

		private void frmCobranzaAsignacionCartera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				dr.Cells["Asignado"].Value = this.chkLocal.Checked;
			}
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			int iSeleccionar = 0;
			if( (int)this.txtSeleccionar.Value < 2) return;
			if(this.uGridDocumentos.Rows.Count > 0)
			{
				int i = 0 ;
				iSeleccionar = Convert.ToInt32 ((int)this.uGridDocumentos.Rows.Count / (int)this.txtSeleccionar.Value);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				{
					dr.Cells["Asignado"].Value = true;
					i = i + 1;
					if( i == iSeleccionar )
					{
						return;
					}
				}
			}
		}

		private void btnVerPagos_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtIPagos.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIPagos.Focus();
				return;
			}
     
				if (this.dtFPagos.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFPagos.Focus();
					return;
				}

				if (Convert.ToDateTime(this.dtIPagos.Value) > Convert.ToDateTime(this.dtFPagos.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFPagos.Focus();
					return;
				}
	
			#endregion Valida Fechas
				if (!Validacion.vbComboVacio(this.cmbEmpresa, "Seleccione la Empresa a Exportar")) return;
			if ((int) this.cmbEmpresa.Value == 2)
			{
 
				string sSQLPa = string.Format("Exec AsignacionCarteraRecoverPagos '{0}', '{1}'", 
					Convert.ToDateTime(this.dtIPagos.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFPagos.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGriPagos);
				this.lblContador.Text = this.uGriPagos.Rows.Count + " Registros Encontrados";
			}
			if ((int) this.cmbEmpresa.Value == 3)
			{
 
				string sSQLPa = string.Format("Exec AsignacionCarteraSiCobraPagos '{0}', '{1}'", 
					Convert.ToDateTime(this.dtIPagos.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFPagos.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGriPagos);
				this.lblContador.Text = this.uGriPagos.Rows.Count + " Registros Encontrados";
			}

		}

		private void btnExportarP_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGriPagos);
		}

		private void btnPagos_Click(object sender, System.EventArgs e)
		{
			int  i=0;
			#region Valida Fechas
			if (this.dtIPagos.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIPagos.Focus();
				return;
			}
     
			if (this.dtFPagos.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFPagos.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtIPagos.Value) > Convert.ToDateTime(this.dtFPagos.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFPagos.Focus();
				return;
			}
	
			#endregion Valida Fechas
			if (!Validacion.vbComboVacio(this.cmbEmpresa, "Seleccione la Empresa a Exportar")) return;
			string sArchivos = "";
			if ((int) this.cmbEmpresa.Value == 2)
			{
				#region Productos 

				SaveFileDialog sArchivoPagos = new SaveFileDialog();

				sArchivoPagos.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoPagos.FileName = "Recover Pagos.txt";
				sArchivoPagos.OverwritePrompt = true;
				sArchivoPagos.Title = "Genera Archivo Pagos Recover";
				                                
				if (sArchivoPagos.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoPagos.FileName, true))
					{
						#region Normales
				
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraRecoverPagos '{0}', '{1}'", 
							Convert.ToDateTime(this.dtIPagos.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFPagos.Value).ToString("yyyyMMdd"))).Rows)
						{
							try
							{
								
								string iLinea = 
									dr["Tipo_Documento"].ToString() 
									+ "|"	+ dr["Tipo_Credito"].ToString() 
									+ "|" + dr["Numero_Identificacion"].ToString() 
									+ "|" + dr["Numero_Operacion"].ToString() 
									+ "|" + dr["OperacionUnica"].ToString() 
									+ "|" + dr["Fecha_Vencimiento"].ToString() 
									+ "|" + dr["Fecha_Pago"].ToString()  
									+ "|"	+ dr["NumeroCuota"].ToString() 
									+ "|" + dr["TotalCobrado"].ToString() 
									+ "|" + dr["Valor_Capital_Cobrado"].ToString() 
									+ "|" + dr["Interes"].ToString() 
									+ "|" + dr["Interes_Mora"].ToString()  
									+ "|"	+ dr["Valor_Honorario_Cobrado"].ToString() 
									+ "|" + dr["Pago_Voluntario"].ToString() 
									+ "|" + dr["Valor_Capital_Condonado"].ToString() 
									+ "|" + dr["Interes_Capital_Condonado"].ToString() 
									+ "|" + dr["Interes_Mora_Condonado"].ToString()  
									+ "|"	+ dr["Valor_Honorario_Condonado"].ToString() 
									+ "|" + dr["Forma_Pago"].ToString() 
									+ "|" + dr["Tipo_Pago"].ToString() 
									+ "|" + dr["Valor_Administrativo"].ToString() 
									;
							
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						#endregion Normales
					}
					if (i > 0) sArchivos = sArchivos + " Productos";
					MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					#endregion Pagos
					
				}

				#endregion Productos 
			}

			if ((int) this.cmbEmpresa.Value == 3)
			{
				#region Productos 

				SaveFileDialog sArchivoPagos = new SaveFileDialog();

				sArchivoPagos.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivoPagos.FileName = "SiCobra Pagos.txt";
				sArchivoPagos.OverwritePrompt = true;
				sArchivoPagos.Title = "Genera Archivo Pagos SiCobra";
				                                
				if (sArchivoPagos.ShowDialog() == DialogResult.OK)
				{					
					#region Pagos
					using (StreamWriter file = new StreamWriter(sArchivoPagos.FileName, true))
					{
						#region Normales
				
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec AsignacionCarteraSiCobraPagos '{0}', '{1}'", 
							Convert.ToDateTime(this.dtIPagos.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFPagos.Value).ToString("yyyyMMdd"))).Rows)
						{
							try
							{
								
								string iLinea = 
								dr["NumeroOperacion"].ToString() 	
								+ "|"	+ dr["Numero_Cuota"].ToString() 	
								+ "|"	+ dr["TipoPago"].ToString() 	
								+ "|"	+ dr["CubrioCuota"].ToString() 	
								+ "|"	+ dr["Pagoefectvio"].ToString() 	
								+ "|"	+ dr["PagoCheque"].ToString() 	
								+ "|"	+ dr["PagoDocumentos"].ToString() 	
								+ "|"	+ dr["Fecha"].ToString() 	
								+ "|"	+ dr["Bodega"].ToString() 	
								+ "|"	+ dr["CanalPago"].ToString() 	
								+ "|"	+ dr["ComprobanteDePago"].ToString() 	
								+ "|"	+ dr["Gastoscobranza"].ToString() 	
								+ "|"	+ dr["ValorReestructurado"].ToString() 	
								+ "|"	+ dr["ValorPagadocuota"].ToString() 	
								+ "|"	+ dr["TipoCancelacion"].ToString() 	
								+ "|"	+ dr["MoraCondonada"].ToString() 	
								+ "|"	+ dr["GastosCobCondonados"].ToString() 	
								+ "|"	+ dr["MoraPagada"].ToString() 	
								+ "|"	+ dr["GastosCobPagado"].ToString();
								file.WriteLine(iLinea);

								i++;

								//		sArchivos = sArchivos +"Referencias/";
							}						
							catch (Exception exc)
							{
								MessageBox.Show(exc.Message + " Linea : " + i.ToString());
								continue;
							}
						}

						#endregion Normales
					}
					if (i > 0) sArchivos = sArchivos + " Productos";
					MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					#endregion Pagos
					
				}

				#endregion Productos 
			}
		}

		private void btnEmpresa_Click(object sender, System.EventArgs e)
		{
			using (frmEmpresasExternas Externa = new frmEmpresasExternas())
			{
				if (DialogResult.OK == Externa.ShowDialog( ))
				{
					this.cmbEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " select idCobranzaExterna, Empresa from CobranzaExterna");
				}
			}
		}

		private void ultraOptionSet1_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optRegion.Value == 0)
			{
				this.cmbProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " SELECT idProvincia, Nombre FROM Cre_Provincia ");
			}
			if ((int) this.optRegion.Value == 1)
			{
				this.cmbProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " SELECT idProvincia, Nombre FROM Cre_Provincia WHERE idProvincia NOT IN (8,13,12,24,9,7) ");
			}
			if ((int) this.optRegion.Value == 2)
			{
				this.cmbProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " SELECT idProvincia, Nombre FROM Cre_Provincia WHERE idProvincia  IN (8,13,12,24,9,7)");
			}
		}

		private void cmbEntidadFinanciera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.cmbProvincia.Focus();
		}

		private void cmbProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.cmbEmpresa.Focus();
		}

		private void cmbEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.btnVer_Click (sender ,e );
		}
	}
}
