using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteDeUna.
	/// </summary>
	public class frmReporteDeUna : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnProcesar;
		private System.Windows.Forms.Label lblCtaCorriente;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label lblCaducidad;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDeposito;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tabsGeneral;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchersConsulta;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchersGeneral;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRetencion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDeposito;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label lblContadorSeleccionados;
		private System.Windows.Forms.Button btnAsiento;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVouchers;
		private System.Windows.Forms.Button btnImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeUna()
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

		bool bNuevo = false;
		bool bEdicion = false;	
		int iContador = 0;
		bool bActualiza = false;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 8, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 8, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepositoDeUna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeUna));
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridVouchers = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridVouchersConsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridVouchersGeneral = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.lblCtaCorriente = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.lblCaducidad = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.tabsGeneral = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblContadorSeleccionados = new System.Windows.Forms.Label();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersGeneral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tabsGeneral)).BeginInit();
			this.tabsGeneral.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.uGridVouchers);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1070, 179);
			// 
			// uGridVouchers
			// 
			this.uGridVouchers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchers.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchers.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 73;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 20;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.Caption = "Doc";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 53;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Almacén";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 140;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 225;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 381;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 104;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridVouchers.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.Appearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.BandLabelAppearance = appearance5;
			this.uGridVouchers.DisplayLayout.GroupByBox.Prompt = " ";
			appearance6.BackColor = System.Drawing.SystemColors.Control;
			appearance6.BackColor2 = System.Drawing.SystemColors.Control;
			appearance6.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance6.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchers.DisplayLayout.GroupByBox.PromptAppearance = appearance6;
			this.uGridVouchers.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVouchers.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVouchers.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchers.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchers.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridVouchers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchers.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridVouchers.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchers.Location = new System.Drawing.Point(8, 8);
			this.uGridVouchers.Name = "uGridVouchers";
			this.uGridVouchers.Size = new System.Drawing.Size(944, 168);
			this.uGridVouchers.TabIndex = 213;
			this.uGridVouchers.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchers_CellChange);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.uGridVouchersConsulta);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(1070, 179);
			// 
			// uGridVouchersConsulta
			// 
			this.uGridVouchersConsulta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchersConsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchersConsulta.DataSource = this.ultraDataSource3;
			appearance11.BackColor = System.Drawing.Color.White;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.Appearance = appearance11;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 73;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Doc";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 134;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.Caption = "Almacén";
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 176;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Width = 317;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 202;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance12;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 93;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance13;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridVouchersConsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.Appearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.BandLabelAppearance = appearance15;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.Prompt = " ";
			appearance16.BackColor = System.Drawing.SystemColors.Control;
			appearance16.BackColor2 = System.Drawing.SystemColors.Control;
			appearance16.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance16.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersConsulta.DisplayLayout.GroupByBox.PromptAppearance = appearance16;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVouchersConsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchersConsulta.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchersConsulta.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridVouchersConsulta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersConsulta.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersConsulta.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridVouchersConsulta.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchersConsulta.Location = new System.Drawing.Point(8, 8);
			this.uGridVouchersConsulta.Name = "uGridVouchersConsulta";
			this.uGridVouchersConsulta.Size = new System.Drawing.Size(944, 168);
			this.uGridVouchersConsulta.TabIndex = 213;
			this.uGridVouchersConsulta.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVouchersConsulta_CellChange);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.optEstados);
			this.ultraTabPageControl1.Controls.Add(this.uGridVouchersGeneral);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1070, 179);
			// 
			// optEstados
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance21;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance22;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = -1;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Capturados";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Depositados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(336, 24);
			this.optEstados.TabIndex = 219;
			this.optEstados.Text = "Todos";
			// 
			// uGridVouchersGeneral
			// 
			this.uGridVouchersGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVouchersGeneral.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVouchersGeneral.DataSource = this.ultraDataSource4;
			appearance23.BackColor = System.Drawing.Color.White;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.Appearance = appearance23;
			this.uGridVouchersGeneral.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 73;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Deposito";
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "";
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 20;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.Caption = "Estado";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Width = 126;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.Width = 97;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Almacén";
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.Width = 254;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Width = 145;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Width = 130;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance24;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 9;
			ultraGridColumn28.Width = 164;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridVouchersGeneral.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.Appearance = appearance25;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.BandLabelAppearance = appearance26;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.Prompt = " ";
			appearance27.BackColor = System.Drawing.SystemColors.Control;
			appearance27.BackColor2 = System.Drawing.SystemColors.Control;
			appearance27.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance27.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVouchersGeneral.DisplayLayout.GroupByBox.PromptAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.uGridVouchersGeneral.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 7.8F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVouchersGeneral.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.uGridVouchersGeneral.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersGeneral.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVouchersGeneral.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.uGridVouchersGeneral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVouchersGeneral.Location = new System.Drawing.Point(8, 40);
			this.uGridVouchersGeneral.Name = "uGridVouchersGeneral";
			this.uGridVouchersGeneral.Size = new System.Drawing.Size(936, 136);
			this.uGridVouchersGeneral.TabIndex = 213;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
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
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 96);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 986;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 96);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 985;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance32;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(240, 94);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 984;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance33;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 94);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 983;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 131);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 1008;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			this.txtBusqueda.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(72, 130);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(272, 19);
			this.txtBusqueda.TabIndex = 1007;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(528, 128);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1006;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(848, 16);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 1010;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(648, 95);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1012;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-64, 160);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1152, 8);
			this.groupBox1.TabIndex = 1046;
			this.groupBox1.TabStop = false;
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnProcesar.Enabled = false;
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnProcesar.Location = new System.Drawing.Point(848, 48);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(88, 23);
			this.btnProcesar.TabIndex = 1047;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Visible = false;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// lblCtaCorriente
			// 
			this.lblCtaCorriente.AutoSize = true;
			this.lblCtaCorriente.BackColor = System.Drawing.Color.Transparent;
			this.lblCtaCorriente.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCtaCorriente.Location = new System.Drawing.Point(16, 40);
			this.lblCtaCorriente.Name = "lblCtaCorriente";
			this.lblCtaCorriente.Size = new System.Drawing.Size(44, 16);
			this.lblCtaCorriente.TabIndex = 1054;
			this.lblCtaCorriente.Text = "Numero";
			this.lblCtaCorriente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblNumero.Location = new System.Drawing.Point(248, 40);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(53, 16);
			this.lblNumero.TabIndex = 1053;
			this.lblNumero.Text = "Concepto";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCaducidad
			// 
			this.lblCaducidad.AutoSize = true;
			this.lblCaducidad.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblCaducidad.Location = new System.Drawing.Point(16, 8);
			this.lblCaducidad.Name = "lblCaducidad";
			this.lblCaducidad.Size = new System.Drawing.Size(36, 16);
			this.lblCaducidad.TabIndex = 1052;
			this.lblCaducidad.Text = "Fecha";
			this.lblCaducidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConcepto
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConcepto.Appearance = appearance34;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(320, 40);
			this.txtConcepto.MaxLength = 200;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(250, 21);
			this.txtConcepto.TabIndex = 1050;
			// 
			// txtDeposito
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeposito.Appearance = appearance35;
			this.txtDeposito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDeposito.Enabled = false;
			this.txtDeposito.Location = new System.Drawing.Point(72, 40);
			this.txtDeposito.MaxLength = 15;
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.Size = new System.Drawing.Size(150, 21);
			this.txtDeposito.TabIndex = 1051;
			// 
			// dtFecha
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance36;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton3.Caption = "Today";
			dateButton3.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1049;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 400);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1056;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(96, 400);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1061;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(182, 400);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 1060;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(351, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 1057;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(523, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1058;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(603, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1059;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// tabsGeneral
			// 
			this.tabsGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.tabsGeneral.Appearance = appearance37;
			this.tabsGeneral.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl9);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl10);
			this.tabsGeneral.Controls.Add(this.ultraTabPageControl1);
			this.tabsGeneral.Location = new System.Drawing.Point(0, 184);
			this.tabsGeneral.Name = "tabsGeneral";
			this.tabsGeneral.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tabsGeneral.Size = new System.Drawing.Size(1072, 200);
			this.tabsGeneral.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.tabsGeneral.TabIndex = 1062;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Capturar";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Capturados";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Consultar";
			this.tabsGeneral.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3});
			this.tabsGeneral.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.tabsGeneral.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.tabsGeneral_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1070, 179);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-112, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1088, 8);
			this.groupBox2.TabIndex = 1063;
			this.groupBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Location = new System.Drawing.Point(792, 403);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 1065;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance38;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(840, 401);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(88, 21);
			this.txtTotal.TabIndex = 1064;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Location = new System.Drawing.Point(960, 400);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 1069;
			this.txtEstado.Visible = false;
			// 
			// txtIdRetencion
			// 
			this.txtIdRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRetencion.Location = new System.Drawing.Point(1008, 400);
			this.txtIdRetencion.Name = "txtIdRetencion";
			this.txtIdRetencion.PromptChar = ' ';
			this.txtIdRetencion.Size = new System.Drawing.Size(16, 21);
			this.txtIdRetencion.TabIndex = 1068;
			this.txtIdRetencion.Visible = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(976, 400);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 1067;
			this.txtIdAsiento.Visible = false;
			// 
			// txtIdDeposito
			// 
			this.txtIdDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDeposito.Location = new System.Drawing.Point(992, 400);
			this.txtIdDeposito.Name = "txtIdDeposito";
			this.txtIdDeposito.PromptChar = ' ';
			this.txtIdDeposito.Size = new System.Drawing.Size(16, 21);
			this.txtIdDeposito.TabIndex = 1066;
			this.txtIdDeposito.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Location = new System.Drawing.Point(360, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 1071;
			this.label2.Text = "Local";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbLocal
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance39;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(416, 94);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(200, 21);
			this.cmbLocal.TabIndex = 1070;
			this.cmbLocal.ValueMember = "Bodega";
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(720, 16);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 1072;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(688, 401);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 1073;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContadorSeleccionados
			// 
			this.lblContadorSeleccionados.AutoSize = true;
			this.lblContadorSeleccionados.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorSeleccionados.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorSeleccionados.Location = new System.Drawing.Point(648, 130);
			this.lblContadorSeleccionados.Name = "lblContadorSeleccionados";
			this.lblContadorSeleccionados.Size = new System.Drawing.Size(0, 18);
			this.lblContadorSeleccionados.TabIndex = 1074;
			this.lblContadorSeleccionados.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.Image = ((System.Drawing.Image)(resources.GetObject("btnAsiento.Image")));
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.Location = new System.Drawing.Point(435, 400);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 1075;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(266, 400);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 1076;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmReporteDeUna
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(952, 438);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.lblContadorSeleccionados);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCtaCorriente);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblCaducidad);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdRetencion);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdDeposito);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.tabsGeneral);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmReporteDeUna";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Depositos DeUna";
			this.Load += new System.EventHandler(this.frmReporteDeUna_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchers)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVouchersGeneral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tabsGeneral)).EndInit();
			this.tabsGeneral.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Cursor  = Cursors.WaitCursor;					
						
			if (bEdicion && bActualiza)
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;

				if (!Validacion.vbTexto(this.txtDeposito, 5, 20, "Numero de deposito")) return;

				if (!Validacion.vbTexto(this.txtConcepto, 5, 100, "Concepto")) return;			

				this.GuardaEncabezado();

				bActualiza = false;
			}

			if (this.tabsGeneral.Tabs[0].Active && (int)this.txtIdDeposito.Value > 0)	this.ConsultaPendientes();
			if (this.tabsGeneral.Tabs[1].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaCapturados();
			if (this.tabsGeneral.Tabs[2].Active) this.ConsultaGeneral();

			this.Cursor  = Cursors.Default;			

			#region anterior
			//			if (this.dtDesde.Value == System.DBNull.Value)
			//			{
			//				MessageBox.Show("Seleccione una fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.dtDesde.Focus();
			//				return;
			//			}
			//
			//			if (this.dtHasta.Value == System.DBNull.Value)
			//			{
			//				MessageBox.Show("Seleccione una fecha final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.dtHasta.Focus();
			//				return;
			//			}
			//
			//			string sSQL = string.Format("Exec ReportePagosDeUna '{0}', '{1}', '{2}',{3}",
			//				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString().Trim(),(int)this.optRegiones.Value);
			//			FuncionesProcedimientos.dsGeneral(sSQL, this.gridPagos);
			//
			//			this.lblContador.Text = this.gridPagos.Rows.Count + " REGISTROS ENCONTRADOS";

			#endregion anterior
		}

		private void GuardaEncabezado()
		{			
			string sSQLDep = string.Format("Exec GrabaEncabezadoDeUna {0}, '{1}', '{2}', '{3}'", 
				(int)this.txtIdDeposito.Value, this.txtDeposito.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
				this.txtConcepto.Text.ToString());
			this.txtIdDeposito.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLDep, true);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmReporteDeUna_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "111721");
			
			if (!Funcion.Permiso("2152", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Reporte De Una ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");	

			#region Fecha
			this.dtDesde.Value = "01/01/2024";
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde.CalendarInfo.MinDate = (DateTime)dtDesde.Value;
			this.dtHasta.CalendarInfo.MinDate = (DateTime)dtDesde.Value;
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;	

			#endregion Fecha

			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnExcel.Enabled = true;
		}

		private void uGridDocumentos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.tabsGeneral.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchers);
			if (this.tabsGeneral.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchersConsulta);
			if (this.tabsGeneral.Tabs[2].Selected) FuncionesProcedimientos.ExportaExcel(this.uGridVouchersGeneral);
		}

		private void gridPagos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region anterior
			//			if (e.Cell.Column.ToString() == "EstadoBit")
			//			{
			//				this.gridPagos.UpdateData(); 
			//				bool estadoActual = Convert.ToBoolean(e.Cell.Value);
			//				int nuevoEstado = estadoActual ? 1 : 0;
			//
			//				int idFormaPago = Convert.ToInt32(e.Cell.Row.Cells["idDetFormaPago"].Value);
			//
			//				string estadoOriginal = Convert.ToString(e.Cell.Row.Cells["Estado"].Value).ToUpper().Trim();
			//
			//				if (estadoOriginal == "DEPOSITADO") 
			//				{
			//					MessageBox.Show("No se puede desmarcar un pago que está Depositado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			//
			//					e.Cell.Row.Cells["EstadoBit"].Value = true;
			//					return;
			//				}
			//
			//				if (idFormaPago > 0)
			//				{
			//					string sSQL = string.Format("Exec ActualizaEstadoDeUna {0}, {1}", idFormaPago, nuevoEstado);
			//					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			//
			//					e.Cell.Row.Cells["EstadoBit"].Value = (nuevoEstado == 1);
			//
			//					this.gridPagos.Refresh();
			//					this.gridPagos.Rows.Refresh(Infragistics.Win.UltraWinGrid.RefreshRow.RefreshDisplay);
			//				}
			//				else
			//				{
			//					MessageBox.Show("Error: El ID del pago no es válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//				}
			//
			//				this.btnVer_Click(this, EventArgs.Empty);
			//			}
			#endregion anterior
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{

		}

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{	
			#region anterior 
			//			decimal Total = 0.00m;
			//			this.txtidAsiento.Value = 0;
			//			if (gridPagos.Rows.Count == 0)
			//			{
			//				MessageBox.Show("No Hay Datos Seleccionados Para Procesar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			//				return;
			//			}
			//
			//			if (DialogResult.Yes == MessageBox.Show("¿Está seguro de procesar los pagos seleccionados?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			//			{		
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
			//				{
			//					Total += Convert.ToDecimal(dr.Cells["Valor"].Value);
			//				}
			//				string sSQL = string.Format("Exec AsientoLiquidacionDEUNA {0}, {1},", (int)this.txtidAsiento.Value, Total);
			//				this.txtidAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			//
			//				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
			//				{
			//					int idFormaPago = Convert.ToInt32(dr.Cells["idDetFormaPago"].Value);
			//					string sSQLUpdate = string.Format("Exec ActualizaEstadoDeUna {0}, {1}", idFormaPago, 2);
			//					Funcion.EjecutaSQL(cdsSeteoF, sSQLUpdate);
			//				}
			//
			//				#region Impresion
			//				string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();				
			//				Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
			//				miRepor.MdiParent = this.MdiParent;
			//				miRepor.Titulo("Asiento");			
			//				miRepor.Show();	
			//				#endregion Impresion
			//}

			#endregion anterior
		}

		private void txtidAsiento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
			DateTime dtFechaVal1 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT DATEADD(DAY,3,EOMONTH('{0}',-1))",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));
			DateTime dtFechaVal2 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT EOMONTH('{0}',-1)",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));	

			if (miAcceso.CambiarFecha)
			{
				this.dtFecha.Enabled = true;
				this.dtFecha.CalendarInfo.MinDate = DateTime.MinValue; 
				this.dtFecha.CalendarInfo.MaxDate = DateTime.MaxValue;
			}
			else
			{
				this.dtFecha.Enabled = true;
				this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-3);
				this.dtFecha.CalendarInfo.MaxDate = dtHoy;
			}

			this.dtFecha.Enabled = true;

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			this.txtDeposito.Enabled = true;
			this.txtConcepto.Enabled = true;
						
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchersConsulta);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.GuardaEncabezado();
			
			bActualiza = true;

			bNuevo = true;
			bEdicion = true;

			this.dtFecha.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bNuevo)
			{
				if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de cancelar esta captura con {0} documentos seleccionados?", iContador), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CancelaDepositoDeUna {0}", (int)this.txtIdDeposito.Value));
				else return;		
			}
	
			bNuevo = false;
			bEdicion = false;

			this.txtIdDeposito.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtIdRetencion.Value = 0;

			this.dtFecha.Value = System.DBNull.Value;			
			this.txtDeposito.Text = "";
			this.txtConcepto.Text = "";

			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.lblContador.Text = "";
			this.lblContadorSeleccionados.Text = "";

			this.txtBusqueda.Text = "";		

			this.dtFecha.Enabled = false;
			this.txtDeposito.Enabled = false;
			this.txtConcepto.Enabled = false;
			
			string sSQL = "Select dfp.idDetFormaPago, dfp.idCompra, 0 As idDepositoDeUna, 0 As Estado, '' As Origen, '' As Almacen, '' As Numero, Null  As Fecha, dfp.Valor, 0 As Bodega, dfp.idFormaPago From DetalleFormaPago dfp Where idDetFormaPago = 0";

			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchers);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridVouchersConsulta);

			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnAsiento.Enabled = false;			
			this.btnCancelar.Enabled = false;
			this.btnImprimir.Enabled = false;
			iContador = 0;
			bActualiza = false;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{		
			using (frmConsultaDepositosDeUna Consulta = new frmConsultaDepositosDeUna())
				if (DialogResult.OK == Consulta.ShowDialog())					
					this.Consultar((int)Consulta.uGridEncabezado.ActiveRow.Cells["idDepositoDeUna"].Value);		
		}
		private void Consultar (int idDepositoDeUna)
		{
			string sSQL = string.Format("Exec DTConsultaIndividualDeUna {0}", idDepositoDeUna);
			SqlDataReader drDeposito = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drDeposito.Read();
			if (drDeposito.HasRows)
			{				
				this.txtIdDeposito.Value = drDeposito.GetInt32(0);
				this.txtIdAsiento.Value = drDeposito.GetInt32(1);
				this.dtFecha.Value = drDeposito.GetDateTime(2);
				this.txtDeposito.Value = drDeposito.GetString(3);	
				this.txtConcepto.Text = drDeposito.GetString(4);
				this.txtTotal.Value = drDeposito.GetDecimal(5);
				this.txtEstado.Value = drDeposito.GetInt32(6);
				this.lblUsuario.Text = drDeposito.GetString(7);
			}
			drDeposito.Close();
						
			if ((int)this.txtEstado.Value == 0) 
			{					
				this.lblEstado.Text = "DEPOSITADO";
				
				this.btnEditar.Enabled = true;
				this.btnAsiento.Enabled = true;		
				this.btnImprimir.Enabled = true;	

				this.tabsGeneral.Tabs[1].Active = true;
								
				this.ConsultaPendientes();
				this.ConsultaCapturados();
			}

			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "ANULADO";

			this.btnCancelar.Enabled = true;				
		}

		private void ConsultaPendientes()
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			int iBodega = 0;
			if (this.cmbLocal.ActiveRow != null) iBodega = (int)this.cmbLocal.Value;			

			string sSQL = string.Format("Exec ListaPendientesDeDepositarDeUna '{0}', '{1}', {2}, '{3}', {4}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iBodega, this.txtBusqueda.Text.ToString(), (int)this.txtIdDeposito.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridVouchers);

			this.lblContador.Text = this.uGridVouchers.Rows.Count + " REGISTROS ENCONTRADOS";	
		}
		private void ConsultaCapturados()
		{			
			string sSQLc = string.Format("Exec ListaDeUnaRegistrados {0}", (int)this.txtIdDeposito.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQLc, this.uGridVouchersConsulta);			

			this.lblContador.Text = this.uGridVouchers.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtIdDeposito.Value, 8, cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;

			this.dtFecha.Enabled = true;
			this.txtDeposito.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnAsiento.Enabled = false;			
			this.btnCancelar.Enabled = true;
			this.btnImprimir.Enabled = false;

		
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchers);			
			FuncionesProcedimientos.EstadoGrids(true, this.uGridVouchersConsulta);
			
			this.ConsultaPendientes();
			this.ConsultaCapturados();

			bNuevo = false;
			bEdicion = true;

			this.dtFecha.Focus();
		}
		private void ConsultaGeneral()
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			int iBodega = 0;

			if (this.cmbLocal.ActiveRow != null) iBodega = (int)this.cmbLocal.Value;			

			string sSQL = string.Format("Exec ListaDepositadosDeUna '{0}', '{1}', {2}, '{3}', {4}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iBodega, this.txtBusqueda.Text.ToString(), (int)this.optEstados.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridVouchersGeneral);		

			this.lblContador.Text = this.uGridVouchersGeneral.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void Suma()
		{
			this.txtTotal.Value = 0;

			this.lblContadorSeleccionados.Text = "";			
			decimal dSuma = 0;

			string sSQL = string.Format("Exec ContadorSumatoriaDeUna {0}, {1}", (int)this.txtIdDeposito.Value, MenuLatinium.IdUsuario);
			SqlDataReader drSuma = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drSuma.Read();
			if (drSuma.HasRows)
			{
				iContador = (int)drSuma.GetValue(0);
				dSuma = Convert.ToDecimal(drSuma.GetValue(1));
			}
			drSuma.Close();

			this.lblContadorSeleccionados.Text = iContador.ToString() + " Documentos Capturados";

			this.txtTotal.Value = dSuma;
		}


		private void tabsGeneral_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
			if (this.tabsGeneral.Tabs[0].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaPendientes();
			if (this.tabsGeneral.Tabs[1].Active && (int)this.txtIdDeposito.Value > 0) this.ConsultaCapturados();
			if (this.tabsGeneral.Tabs[2].Active)	this.ConsultaGeneral();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.tabsGeneral.Tabs[1].Selected = true;

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha del deposito", true, "Bancos", cdsSeteoF)) return;				

			if (!Validacion.vbTexto(this.txtDeposito, 5, 20, "Numero de deposito")) return;

			if (!Validacion.vbTexto(this.txtConcepto, 5, 100, "Concepto")) return;
			
			if (this.uGridVouchersConsulta.Rows.Count == 0)
			{
				MessageBox.Show("Seleccione el o los Depositos De Una que va a Procesar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			if (DialogResult.Yes == MessageBox.Show("ANTES DE GENERAR EL DEPOSITO. \n\nVERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n\n¿Esta seguro de Generar el Deposito?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{				
				#region Cálculo de valores
				decimal dTotal = 0.00m;
				DateTime dtFechaDeposito = (DateTime)this.dtFecha.Value;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchersConsulta.Rows.All)
				{
					dTotal = dTotal + Convert.ToDecimal(dr.Cells["Total"].Value);
				}
				#endregion Cálculo de valores
			
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 120;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						string sSQL = string.Format("Exec GrabaEncabezadoDeUna {0}, '{1}', '{2}', '{3}', '{4}'", 
							(int)this.txtIdDeposito.Value,
							this.txtDeposito.Text.ToString(), 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							this.txtConcepto.Text.ToString(), 
							Convert.ToDecimal(this.txtTotal.Value));
						oCmdActualiza.CommandText = sSQL;
						this.txtIdDeposito.Value = (int)oCmdActualiza.ExecuteScalar();
						
						#region Genera Asiento
						string sNota = this.txtConcepto.Text.ToString();
						string sNumero = this.txtDeposito.Text.ToString();
						sSQL = string.Format("Exec AsientoLiquidacionDEUNA {0},'{1}','{2}','{3}','{4}'", 
							(int)this.txtIdAsiento.Value,
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
							dTotal,
							sNota,
							sNumero);
						oCmdActualiza.CommandText = sSQL;
						this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

						oCmdActualiza.CommandText = string.Format("Update DepositoDeuna Set idAsiento = {0} Where idDepositoDeUna = {1}", (int)this.txtIdAsiento.Value, (int)this.txtIdDeposito.Value);
						oCmdActualiza.ExecuteNonQuery();
						#endregion Genera Asiento

						#region Actualiza Detalle de la forma pago
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVouchersConsulta.Rows.All)
						{
							if ((int)dr.Cells["Estado"].Value == 1)
							{
								oCmdActualiza.CommandText = string.Format("Update DetalleFormaPago Set Estado = 2 Where idDetFormaPago = {0}", (int)dr.Cells["idDetFormaPago"].Value);
								oCmdActualiza.ExecuteNonQuery();							
							}
						}
						#endregion Actualiza Detalle de la forma pago					

						oTransaction.Commit();

						MessageBox.Show("Depósito registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
						
						
						#region Imprime reporte
						
						string strHostName = string.Empty;
						strHostName = System.Environment.MachineName;
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE DepositoDeUna set iNumImp = isnull ( iNumImp, 0) + 1 where idAsiento = {0}  ", (int)this.txtIdAsiento.Value));										

						string sReporte = "DepositoDeUna.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						if (MenuLatinium.stDirFacturacion == "DENTALIS")
						{
							crConnectionInfo.ServerName = strServidor; //"172.16.1.2"
							crConnectionInfo.DatabaseName = strBaseDato; //"DENTALIS"
							crConnectionInfo.UserID = "sa"; 
							crConnectionInfo.Password = "DentServ08";
						}
						else
						{
							crConnectionInfo.ServerName = strServidor;
							crConnectionInfo.DatabaseName = strBaseDato; 
							crConnectionInfo.UserID = "Infoelect"; 
							crConnectionInfo.Password = "Bl45o6$9"; 
						}
					}
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
						Tables crTables = crDatabase.Tables; 
						Table crTable;

						for (int i=0; i< crTables.Count; i++)
						{
						{
							crTable = crTables [i]; 
							crTableLogOnInfo = crTable.LogOnInfo; 
							crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
							crTable.ApplyLogOnInfo(crTableLogOnInfo);
						}
						}
					
						if (MenuLatinium.stDirFacturacion == "POINT" || MenuLatinium.stDirFacturacion == "CREDISOLUTIONS" || MenuLatinium.stDirFacturacion == "DESAPOINT")
						{
							Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
							Report1.SetParameterValue("@idAsiento", (int)this.txtIdAsiento.Value);
							Report1.SetParameterValue("@idAsientoDet", (int)this.txtIdAsiento.Value);
							Report1.SetParameterValue("@EstacionImp", "@EstacionImp");
							Report1.SetParameterValue("@UsuarioImp", "@UsuarioImp");
							Report1.SetParameterValue("@idDepositoTarjeta", (int)this.txtIdDeposito.Value);
				
							Report1.PrintToPrinter(1, false, 0, 0);
						}
						else
						{
							int idDeposito = (int) this.txtIdDeposito.Value;
							ParameterFields paramFields = new ParameterFields ();
							
							ParameterField pfidAsiento = new ParameterField ();
							ParameterDiscreteValue ValidAsiento = new ParameterDiscreteValue ();
							pfidAsiento.ParameterFieldName = "@idAsiento";
							ValidAsiento.Value = (int)this.txtIdAsiento.Value;
							pfidAsiento.CurrentValues.Add (ValidAsiento);
							paramFields.Add (pfidAsiento);
							
							ParameterField pfidAsientoDet = new ParameterField ();
							ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
							pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
							ValIdAsientoDet.Value = (int)this.txtIdAsiento.Value;
							pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
							paramFields.Add (pfidAsientoDet);
			
							ParameterField paramFieldEs = new ParameterField ();
							ParameterDiscreteValue discreteValEs = new ParameterDiscreteValue ();
							paramFieldEs.ParameterFieldName = "@EstacionImp";
							discreteValEs.Value = strHostName;
							paramFieldEs.CurrentValues.Add (discreteValEs);
							paramFields.Add (paramFieldEs);			
			
							ParameterField paramFieldUs = new ParameterField ();
							ParameterDiscreteValue discreteValUs = new ParameterDiscreteValue ();
							paramFieldUs.ParameterFieldName = "@UsuarioImp";
							discreteValUs.Value = MenuLatinium.stUsuario;
							paramFieldUs.CurrentValues.Add (discreteValUs);
							paramFields.Add (paramFieldUs);			
			
							ParameterField pfidIdDeposito = new ParameterField ();
							ParameterDiscreteValue ValIdIdDeposito = new ParameterDiscreteValue ();
							pfidIdDeposito.ParameterFieldName = "@idDepositoDeUna";
							ValIdIdDeposito.Value = 		idDeposito;
							pfidIdDeposito.CurrentValues.Add (ValIdIdDeposito);
							paramFields.Add (pfidIdDeposito);
											
							Reporte miRepor = new Reporte("DepositoDeUna.rpt", "");
							miRepor.MdiParent = this.MdiParent;
							miRepor.Titulo("Depósito de Una");
							miRepor.crVista.ParameterFieldInfo = paramFields;
							miRepor.Show();
						}
						#endregion Imprime reporte

						#region Controles
						this.dtFecha.Enabled = false;
						this.txtDeposito.Enabled = false;
						this.txtConcepto.Enabled = false;		
						this.btnImprimir.Enabled = true;

						FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchers);
						FuncionesProcedimientos.EstadoGrids(false, this.uGridVouchersConsulta);

						string strFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.Value.ToString();
						string strReporte = "Asientos.rpt";

						if (MessageBox.Show("¿Desea visualizar el asiento generado?", "Impresión de Asiento", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
						{
							Reporte miReporte = new Reporte(strReporte, strFiltro);
							miReporte.MdiParent = MdiParent;
							miReporte.Show();
						}
					

						this.btnNuevo.Enabled = true;
						this.btnConsultar.Enabled = true;
						this.btnEditar.Enabled = true;
						this.btnGuardar.Enabled = false;			
						this.btnAsiento.Enabled = true;							

						bEdicion = false;
						bNuevo = false;
						#endregion Controles
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					}
					finally
					{
						oConexion.Close();
					}
				}			
			}	
		}

		private void uGridVouchersConsulta_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Estado")
			{
				this.uGridVouchers.UpdateData();
				
				string sSQL = string.Format("Exec CVValidaRegistroCapturadoDeUna {0}, {1}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				
				if (sMensaje.Length > 0)
				{
					e.Cell.Row.Cells["Estado"].Value = false;
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				sSQL = string.Format("Exec CVActualizaCapturadoDeUna {0}, {1}, {2}", 
					(bool)e.Cell.Row.Cells["Estado"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);									

				Suma();				
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			frmAsientos Asientos = new frmAsientos(((int)this.txtIdAsiento.Value));
			Asientos.MdiParent = MdiParent;
			Asientos.Show();	
		}

		private void uGridVouchers_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridVouchers_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Estado")
			{
				this.uGridVouchers.UpdateData();
				
				string sSQL = string.Format("Exec CVValidaRegistroCapturadoDeUna {0}, {1}", (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				
				if (sMensaje.Length > 0)
				{
					e.Cell.Row.Cells["Estado"].Value = false;
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				sSQL = string.Format("Exec CVActualizaCapturadoDeUna {0}, {1}, {2}", 
					(bool)e.Cell.Row.Cells["Estado"].Value, (int)e.Cell.Row.Cells["idDetFormaPago"].Value, (int)this.txtIdDeposito.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);									

				Suma();				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
		
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No Existe Deposito para Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			int idDeposito = 0;
			string strHostName = string.Empty;
			strHostName = System.Environment.MachineName;
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE DepositoDeUna set iNumImp = isnull ( iNumImp, 0) + 1 where idAsiento = {0}  ", (int)this.txtIdAsiento.Value));						

			idDeposito = (int) this.txtIdDeposito.Value;
		
			#region Imprime reporte
			ParameterFields paramFields = new ParameterFields ();
							
			ParameterField pfidAsiento = new ParameterField ();
			ParameterDiscreteValue ValidAsiento = new ParameterDiscreteValue ();
			pfidAsiento.ParameterFieldName = "@idAsiento";
			ValidAsiento.Value = (int)this.txtIdAsiento.Value;
			pfidAsiento.CurrentValues.Add (ValidAsiento);
			paramFields.Add (pfidAsiento);
							
			ParameterField pfidAsientoDet = new ParameterField ();
			ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
			pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
			ValIdAsientoDet.Value = (int)this.txtIdAsiento.Value;
			pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
			paramFields.Add (pfidAsientoDet);
			
			ParameterField paramFieldEs = new ParameterField ();
			ParameterDiscreteValue discreteValEs = new ParameterDiscreteValue ();
			paramFieldEs.ParameterFieldName = "@EstacionImp";
			discreteValEs.Value = strHostName;
			paramFieldEs.CurrentValues.Add (discreteValEs);
			paramFields.Add (paramFieldEs);
			
			
			ParameterField paramFieldUs = new ParameterField ();
			ParameterDiscreteValue discreteValUs = new ParameterDiscreteValue ();
			paramFieldUs.ParameterFieldName = "@UsuarioImp";
			discreteValUs.Value = MenuLatinium.stUsuario;
			paramFieldUs.CurrentValues.Add (discreteValUs);
			paramFields.Add (paramFieldUs);
			
			
			ParameterField pfidIdDeposito = new ParameterField ();
			ParameterDiscreteValue ValIdIdDeposito = new ParameterDiscreteValue ();
			pfidIdDeposito.ParameterFieldName = "@idDepositoDeUna";
			ValIdIdDeposito.Value = idDeposito;
			pfidIdDeposito.CurrentValues.Add (ValIdIdDeposito);
			paramFields.Add (pfidIdDeposito);
	
											
			Reporte miRepor = new Reporte("DepositoDeUna.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Depósito de DeUna");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.Show();	
			#endregion Imprime reporte		
		}
	}
}
