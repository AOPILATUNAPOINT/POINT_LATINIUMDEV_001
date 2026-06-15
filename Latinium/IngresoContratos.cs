using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratos.
	/// </summary>
	public class IngresoContratos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporteCompara;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel20;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel32;
		private Infragistics.Win.Misc.UltraLabel ultraLabel33;
		private System.Windows.Forms.GroupBox groupBox5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel34;
		private Infragistics.Win.Misc.UltraLabel ultraLabel35;
		private C1.Data.C1DataSet cdsMoneda;
		private C1.Data.C1DataView cdvMoneda;
		private Infragistics.Win.Misc.UltraButton btnServicios;
		private Infragistics.Win.Misc.UltraButton btnTabla;
		private Infragistics.Win.Misc.UltraButton btnFactura;
		private C1.Data.C1DataView cdvContratos;
		private C1.Data.C1DataSet cdsContratos;
		private C1.Data.C1DataSet cdsCc;
		private C1.Data.C1DataView cdvCc;
		private C1.Data.C1DataView cdvLocales;
		private C1.Data.C1DataSet cdsLocales;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsActividad;
		private C1.Data.C1DataView cdvActividad;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaF;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaI;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContratante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContratado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaP;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDoc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNum;
		private C1.Data.C1DataSet cdsFinancia;
		private C1.Data.C1DataView cdvFinancia;
		private C1.Data.C1DataView cdvAmortiza;
		private C1.Data.C1DataSet cdsAmortiza;
		private C1.Data.C1DataSet cdsServicio;
		private C1.Data.C1DataView cdvServicio;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbServicio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor idContratos;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.GroupBox groupBox7;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCComercial;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraButton btnLocal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private System.Windows.Forms.GroupBox groupBox1;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTriggerC;
		private Infragistics.Win.Misc.UltraButton btnCliente;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFinquito;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btnVerAdemdums;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAdemdum;
		private Infragistics.Win.Misc.UltraButton btFinanciacion;
	int Valarma = 0;
		public IngresoContratos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();			
			Valarma=1;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}		
		int vIdContrato=0;
		public IngresoContratos(int idContrat)
		{
			InitializeComponent();
			vIdContrato = idContrat;			
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosCentroC", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroComercial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosLocales", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroComercial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUbicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Superficie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActividad");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosServicios", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosFinanciacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFinanciacionS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicio", -1, "cmbServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VTrigger");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorU");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(IngresoContratos));
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btFinanciacion = new Infragistics.Win.Misc.UltraButton();
			this.btnCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbCComercial = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.cdvCc = new C1.Data.C1DataView();
			this.cdsCc = new C1.Data.C1DataSet();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvLocales = new C1.Data.C1DataView();
			this.cdsLocales = new C1.Data.C1DataSet();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.btnLocal = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel35 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel34 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFechaF = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel33 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFechaI = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel32 = new Infragistics.Win.Misc.UltraLabel();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtContratante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtContratado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel20 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFechaP = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtDoc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNum = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.idContratos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTriggerC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbFechaFinquito = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.txtAdemdum = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cmbServicio = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvServicio = new C1.Data.C1DataView();
			this.cdsServicio = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsFinancia = new C1.Data.C1DataSet();
			this.btnFactura = new Infragistics.Win.Misc.UltraButton();
			this.btnServicios = new Infragistics.Win.Misc.UltraButton();
			this.btnTabla = new Infragistics.Win.Misc.UltraButton();
			this.btnVerAdemdums = new Infragistics.Win.Misc.UltraButton();
			this.cdvMoneda = new C1.Data.C1DataView();
			this.cdsMoneda = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdvActividad = new C1.Data.C1DataView();
			this.cdsActividad = new C1.Data.C1DataSet();
			this.cdsAmortiza = new C1.Data.C1DataSet();
			this.cdvContratos = new C1.Data.C1DataView();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtReporteCompara = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdvFinancia = new C1.Data.C1DataView();
			this.cdvAmortiza = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCComercial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.groupBox4.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaI)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtContratante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContratado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNum)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTriggerC)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFinquito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdemdum)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			this.groupBox7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFinancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAmortiza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporteCompara)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdvFinancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAmortiza)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btFinanciacion);
			this.ultraTabPageControl1.Controls.Add(this.btnCliente);
			this.ultraTabPageControl1.Controls.Add(this.cmbCComercial);
			this.ultraTabPageControl1.Controls.Add(this.cmbLocal);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel12);
			this.ultraTabPageControl1.Controls.Add(this.cmbCliente);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel5);
			this.ultraTabPageControl1.Controls.Add(this.btnLocal);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel13);
			this.ultraTabPageControl1.Controls.Add(this.groupBox4);
			this.ultraTabPageControl1.Controls.Add(this.groupBox2);
			this.ultraTabPageControl1.Controls.Add(this.cmbFechaP);
			this.ultraTabPageControl1.Controls.Add(this.txtDoc);
			this.ultraTabPageControl1.Controls.Add(this.txtNum);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel14);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel15);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel16);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel17);
			this.ultraTabPageControl1.Controls.Add(this.idContratos);
			this.ultraTabPageControl1.Controls.Add(this.txtTriggerC);
			this.ultraTabPageControl1.Controls.Add(this.groupBox1);
			this.ultraTabPageControl1.Controls.Add(this.txtAdemdum);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(700, 454);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// btFinanciacion
			// 
			this.btFinanciacion.BackColor = System.Drawing.Color.Transparent;
			this.btFinanciacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFinanciacion.Location = new System.Drawing.Point(224, 40);
			this.btFinanciacion.Name = "btFinanciacion";
			this.btFinanciacion.Size = new System.Drawing.Size(136, 24);
			this.btFinanciacion.TabIndex = 145;
			this.btFinanciacion.Text = "&Financiación Servicios";
			this.btFinanciacion.Click += new System.EventHandler(this.btFinanciacion_Click);
			// 
			// btnCliente
			// 
			this.btnCliente.BackColor = System.Drawing.Color.Transparent;
			this.btnCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnCliente.Location = new System.Drawing.Point(320, 80);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 23);
			this.btnCliente.TabIndex = 142;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// cmbCComercial
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCComercial.Appearance = appearance1;
			this.cmbCComercial.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCComercial.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCComercial.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.idCentroComercial"));
			this.cmbCComercial.DataSource = this.cdvCc;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCComercial.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCComercial.DisplayMember = "Nombre";
			this.cmbCComercial.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCComercial.Enabled = false;
			this.cmbCComercial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCComercial.Location = new System.Drawing.Point(120, 120);
			this.cmbCComercial.Name = "cmbCComercial";
			this.cmbCComercial.Size = new System.Drawing.Size(200, 21);
			this.cmbCComercial.TabIndex = 139;
			this.cmbCComercial.ValueMember = "idCentroComercial";
			this.cmbCComercial.ValueChanged += new System.EventHandler(this.cmbCComercial_ValueChanged);
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsContrato";
			this.cdsContratos.EnforceConstraints = false;
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			this.cdsContratos.BeforeFill += new C1.Data.FillEventHandler(this.cdsContratos_BeforeFill);
			// 
			// cdvCc
			// 
			this.cdvCc.BindingContextCtrl = this;
			this.cdvCc.DataSet = this.cdsCc;
			this.cdvCc.TableName = "";
			this.cdvCc.TableViewName = "ContratosCentroC";
			// 
			// cdsCc
			// 
			this.cdsCc.BindingContextCtrl = this;
			this.cdsCc.DataLibrary = "LibFacturacion";
			this.cdsCc.DataLibraryUrl = "";
			this.cdsCc.DataSetDef = "dsContrato";
			this.cdsCc.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCc.Name = "cdsCc";
			this.cdsCc.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCc.SchemaDef = null;
			this.cdsCc.BeforeFill += new C1.Data.FillEventHandler(this.cdsCc_BeforeFill);
			// 
			// cmbLocal
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance2;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.idLocales"));
			this.cmbLocal.DataSource = this.cdvLocales;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLocal.Enabled = false;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(440, 120);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(200, 21);
			this.cmbLocal.TabIndex = 135;
			this.cmbLocal.ValueMember = "idLocales";
			// 
			// cdvLocales
			// 
			this.cdvLocales.BindingContextCtrl = this;
			this.cdvLocales.DataSet = this.cdsLocales;
			this.cdvLocales.Sort = "idLocales";
			this.cdvLocales.TableName = "";
			this.cdvLocales.TableViewName = "ContratosLocales";
			// 
			// cdsLocales
			// 
			this.cdsLocales.BindingContextCtrl = this;
			this.cdsLocales.DataLibrary = "LibFacturacion";
			this.cdsLocales.DataLibraryUrl = "";
			this.cdsLocales.DataSetDef = "dsContrato";
			this.cdsLocales.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsLocales.Name = "cdsLocales";
			this.cdsLocales.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsLocales.SchemaDef = null;
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel12.Location = new System.Drawing.Point(376, 120);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(32, 16);
			this.ultraLabel12.TabIndex = 137;
			this.ultraLabel12.Text = "Local:";
			// 
			// cmbCliente
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance3;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridBand3.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 262;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 9;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(120, 80);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(200, 21);
			this.cmbCliente.TabIndex = 134;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "Proveedor = 0";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel5.Location = new System.Drawing.Point(32, 80);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(40, 16);
			this.ultraLabel5.TabIndex = 136;
			this.ultraLabel5.Text = "Cliente:";
			// 
			// btnLocal
			// 
			this.btnLocal.BackColor = System.Drawing.Color.Transparent;
			this.btnLocal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnLocal.Location = new System.Drawing.Point(640, 120);
			this.btnLocal.Name = "btnLocal";
			this.btnLocal.Size = new System.Drawing.Size(24, 23);
			this.btnLocal.TabIndex = 138;
			this.btnLocal.Text = "...";
			this.btnLocal.Click += new System.EventHandler(this.btnLocal_Click_1);
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel13.Location = new System.Drawing.Point(32, 128);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(104, 16);
			this.ultraLabel13.TabIndex = 133;
			this.ultraLabel13.Text = "Centro Comercial:";
			// 
			// groupBox4
			// 
			this.groupBox4.BackColor = System.Drawing.Color.Transparent;
			this.groupBox4.Controls.Add(this.groupBox5);
			this.groupBox4.Controls.Add(this.cmbFechaF);
			this.groupBox4.Controls.Add(this.ultraLabel33);
			this.groupBox4.Controls.Add(this.cmbFechaI);
			this.groupBox4.Controls.Add(this.ultraLabel32);
			this.groupBox4.Enabled = false;
			this.groupBox4.Location = new System.Drawing.Point(24, 184);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(392, 128);
			this.groupBox4.TabIndex = 131;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Datos del Contrato";
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.txtMes);
			this.groupBox5.Controls.Add(this.txtAnio);
			this.groupBox5.Controls.Add(this.ultraLabel35);
			this.groupBox5.Controls.Add(this.ultraLabel34);
			this.groupBox5.Location = new System.Drawing.Point(240, 16);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(128, 96);
			this.groupBox5.TabIndex = 130;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Duración";
			// 
			// txtMes
			// 
			this.txtMes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.mes"));
			this.txtMes.Location = new System.Drawing.Point(72, 56);
			this.txtMes.MaskInput = "nnn";
			this.txtMes.MinValue = 0;
			this.txtMes.Name = "txtMes";
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(48, 21);
			this.txtMes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtMes.TabIndex = 112;
			this.txtMes.ValueChanged += new System.EventHandler(this.txtMes_ValueChanged);
			// 
			// txtAnio
			// 
			this.txtAnio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.anio"));
			this.txtAnio.Location = new System.Drawing.Point(72, 24);
			this.txtAnio.MaskInput = "nn";
			this.txtAnio.MinValue = 0;
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(48, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 111;
			this.txtAnio.ValueChanged += new System.EventHandler(this.txtAnio_ValueChanged);
			// 
			// ultraLabel35
			// 
			this.ultraLabel35.Location = new System.Drawing.Point(12, 64);
			this.ultraLabel35.Name = "ultraLabel35";
			this.ultraLabel35.Size = new System.Drawing.Size(52, 16);
			this.ultraLabel35.TabIndex = 110;
			this.ultraLabel35.Text = "Meses";
			// 
			// ultraLabel34
			// 
			this.ultraLabel34.Location = new System.Drawing.Point(16, 32);
			this.ultraLabel34.Name = "ultraLabel34";
			this.ultraLabel34.Size = new System.Drawing.Size(52, 16);
			this.ultraLabel34.TabIndex = 108;
			this.ultraLabel34.Text = "Años";
			// 
			// cmbFechaF
			// 
			this.cmbFechaF.AllowMonthSelection = true;
			this.cmbFechaF.AllowWeekSelection = true;
			this.cmbFechaF.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaF.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.FechaF"));
			dateButton1.Caption = "Today";
			this.cmbFechaF.DateButtons.Add(dateButton1);
			this.cmbFechaF.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaF.Format = "dd/MMM/yyyy";
			this.cmbFechaF.Location = new System.Drawing.Point(72, 80);
			this.cmbFechaF.Name = "cmbFechaF";
			this.cmbFechaF.NonAutoSizeHeight = 23;
			this.cmbFechaF.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaF.SpinButtonsVisible = true;
			this.cmbFechaF.TabIndex = 128;
			this.cmbFechaF.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			this.cmbFechaF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFechaF_KeyPress);
			this.cmbFechaF.TextChanged += new System.EventHandler(this.cmbFechaF_TextChanged);
			// 
			// ultraLabel33
			// 
			this.ultraLabel33.Location = new System.Drawing.Point(8, 88);
			this.ultraLabel33.Name = "ultraLabel33";
			this.ultraLabel33.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel33.TabIndex = 129;
			this.ultraLabel33.Text = "Fecha Fin:";
			// 
			// cmbFechaI
			// 
			this.cmbFechaI.AllowMonthSelection = true;
			this.cmbFechaI.AllowWeekSelection = true;
			this.cmbFechaI.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaI.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.FechaI"));
			dateButton2.Caption = "Today";
			this.cmbFechaI.DateButtons.Add(dateButton2);
			this.cmbFechaI.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaI.Format = "dd/MMM/yyyy";
			this.cmbFechaI.Location = new System.Drawing.Point(72, 56);
			this.cmbFechaI.Name = "cmbFechaI";
			this.cmbFechaI.NonAutoSizeHeight = 23;
			this.cmbFechaI.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaI.SpinButtonsVisible = true;
			this.cmbFechaI.TabIndex = 126;
			this.cmbFechaI.Value = new System.DateTime(2010, 3, 4, 0, 0, 0, 0);
			// 
			// ultraLabel32
			// 
			this.ultraLabel32.Location = new System.Drawing.Point(8, 40);
			this.ultraLabel32.Name = "ultraLabel32";
			this.ultraLabel32.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel32.TabIndex = 127;
			this.ultraLabel32.Text = "Fecha Inicio:";
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.Color.Transparent;
			this.groupBox2.Controls.Add(this.txtContratante);
			this.groupBox2.Controls.Add(this.ultraLabel2);
			this.groupBox2.Controls.Add(this.txtContratado);
			this.groupBox2.Controls.Add(this.ultraLabel20);
			this.groupBox2.Enabled = false;
			this.groupBox2.Location = new System.Drawing.Point(24, 352);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(672, 72);
			this.groupBox2.TabIndex = 129;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Firmas";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// txtContratante
			// 
			this.txtContratante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Contratante"));
			this.txtContratante.Location = new System.Drawing.Point(416, 32);
			this.txtContratante.Name = "txtContratante";
			this.txtContratante.Size = new System.Drawing.Size(248, 21);
			this.txtContratante.TabIndex = 129;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(8, 32);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel2.TabIndex = 128;
			this.ultraLabel2.Text = "Contratista";
			// 
			// txtContratado
			// 
			this.txtContratado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Contratado"));
			this.txtContratado.Location = new System.Drawing.Point(88, 32);
			this.txtContratado.Name = "txtContratado";
			this.txtContratado.Size = new System.Drawing.Size(248, 21);
			this.txtContratado.TabIndex = 127;
			// 
			// ultraLabel20
			// 
			this.ultraLabel20.Location = new System.Drawing.Point(344, 32);
			this.ultraLabel20.Name = "ultraLabel20";
			this.ultraLabel20.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel20.TabIndex = 128;
			this.ultraLabel20.Text = "Contratante";
			// 
			// cmbFechaP
			// 
			this.cmbFechaP.AllowMonthSelection = true;
			this.cmbFechaP.AllowWeekSelection = true;
			this.cmbFechaP.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaP.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.FechaP"));
			dateButton3.Caption = "Today";
			this.cmbFechaP.DateButtons.Add(dateButton3);
			this.cmbFechaP.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaP.Enabled = false;
			this.cmbFechaP.Format = "dd/MMM/yyyy";
			this.cmbFechaP.Location = new System.Drawing.Point(584, 8);
			this.cmbFechaP.Name = "cmbFechaP";
			this.cmbFechaP.NonAutoSizeHeight = 23;
			this.cmbFechaP.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaP.SpinButtonsVisible = true;
			this.cmbFechaP.TabIndex = 3;
			this.cmbFechaP.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			// 
			// txtDoc
			// 
			this.txtDoc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Doc"));
			this.txtDoc.Enabled = false;
			this.txtDoc.Location = new System.Drawing.Point(224, 40);
			this.txtDoc.Name = "txtDoc";
			this.txtDoc.Size = new System.Drawing.Size(72, 21);
			this.txtDoc.TabIndex = 2;
			this.txtDoc.Visible = false;
			// 
			// txtNum
			// 
			this.txtNum.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Codigo"));
			this.txtNum.Enabled = false;
			this.txtNum.Location = new System.Drawing.Point(120, 40);
			this.txtNum.Name = "txtNum";
			this.txtNum.Size = new System.Drawing.Size(80, 21);
			this.txtNum.TabIndex = 1;
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel14.Location = new System.Drawing.Point(480, 16);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel14.TabIndex = 125;
			this.ultraLabel14.Text = "Fecha Proceso:";
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel15.Location = new System.Drawing.Point(32, 48);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel15.TabIndex = 124;
			this.ultraLabel15.Text = "Contrato:";
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel16.Location = new System.Drawing.Point(224, 24);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel16.TabIndex = 123;
			this.ultraLabel16.Text = "Num. Documento:";
			this.ultraLabel16.Visible = false;
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel17.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel17.Location = new System.Drawing.Point(48, 8);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(136, 40);
			this.ultraLabel17.TabIndex = 122;
			this.ultraLabel17.Text = "Ingresos de Contratos";
			// 
			// idContratos
			// 
			this.idContratos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.idContratos"));
			this.idContratos.Enabled = false;
			this.idContratos.Location = new System.Drawing.Point(112, 8);
			this.idContratos.Name = "idContratos";
			this.idContratos.Size = new System.Drawing.Size(72, 21);
			this.idContratos.TabIndex = 132;
			// 
			// txtTriggerC
			// 
			this.txtTriggerC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.VTrigger"));
			this.txtTriggerC.Location = new System.Drawing.Point(72, 24);
			this.txtTriggerC.Name = "txtTriggerC";
			this.txtTriggerC.Size = new System.Drawing.Size(72, 21);
			this.txtTriggerC.TabIndex = 141;
			this.txtTriggerC.Text = "0";
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.Transparent;
			this.groupBox1.Controls.Add(this.cmbFechaFinquito);
			this.groupBox1.Controls.Add(this.txtNotas);
			this.groupBox1.Controls.Add(this.chkCliente);
			this.groupBox1.Enabled = false;
			this.groupBox1.Location = new System.Drawing.Point(424, 184);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(264, 128);
			this.groupBox1.TabIndex = 140;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Finiquito";
			// 
			// cmbFechaFinquito
			// 
			this.cmbFechaFinquito.AllowMonthSelection = true;
			this.cmbFechaFinquito.AllowWeekSelection = true;
			this.cmbFechaFinquito.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaFinquito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.FechaFiniquito"));
			dateButton4.Caption = "Today";
			this.cmbFechaFinquito.DateButtons.Add(dateButton4);
			this.cmbFechaFinquito.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaFinquito.Format = "dd/MMM/yyyy";
			this.cmbFechaFinquito.Location = new System.Drawing.Point(144, 16);
			this.cmbFechaFinquito.Name = "cmbFechaFinquito";
			this.cmbFechaFinquito.NonAutoSizeHeight = 23;
			this.cmbFechaFinquito.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaFinquito.SpinButtonsVisible = true;
			this.cmbFechaFinquito.TabIndex = 126;
			this.cmbFechaFinquito.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance4;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Observacion"));
			this.txtNotas.Location = new System.Drawing.Point(16, 48);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(240, 64);
			this.txtNotas.TabIndex = 83;
			// 
			// chkCliente
			// 
			this.chkCliente.DataBindings.Add(new System.Windows.Forms.Binding("Tag", this.cdsContratos, "Contratos.Finiquito"));
			this.chkCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsContratos, "Contratos.Finiquito"));
			this.chkCliente.Location = new System.Drawing.Point(16, 16);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Terminar Contrato";
			this.chkCliente.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.Transparent, System.Drawing.SystemColors.ControlText);
			this.chkCliente.Size = new System.Drawing.Size(128, 19);
			this.chkCliente.TabIndex = 57;
			this.chkCliente.Validating += new System.ComponentModel.CancelEventHandler(this.chkCliente_Validating);
			// 
			// txtAdemdum
			// 
			this.txtAdemdum.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "Contratos.Ademdum"));
			this.txtAdemdum.Enabled = false;
			this.txtAdemdum.Location = new System.Drawing.Point(48, 16);
			this.txtAdemdum.Name = "txtAdemdum";
			this.txtAdemdum.Size = new System.Drawing.Size(72, 21);
			this.txtAdemdum.TabIndex = 144;
			this.txtAdemdum.Visible = false;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl2.Controls.Add(this.groupBox7);
			this.ultraTabPageControl2.Controls.Add(this.groupBox3);
			this.ultraTabPageControl2.Controls.Add(this.btnFactura);
			this.ultraTabPageControl2.Controls.Add(this.btnServicios);
			this.ultraTabPageControl2.Controls.Add(this.btnTabla);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(700, 454);
			this.ultraTabPageControl2.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl2_Paint);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.BackColor = System.Drawing.Color.Transparent;
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(504, 192);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(88, 24);
			this.ultraButton1.TabIndex = 131;
			this.ultraButton1.Text = "Guardar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// groupBox7
			// 
			this.groupBox7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox7.BackColor = System.Drawing.Color.Transparent;
			this.groupBox7.Controls.Add(this.ultraGrid1);
			this.groupBox7.Location = new System.Drawing.Point(8, 216);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(688, 200);
			this.groupBox7.TabIndex = 130;
			this.groupBox7.TabStop = false;
			this.groupBox7.Text = "Tabla Amortización";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			appearance10.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(688, 184);
			this.ultraGrid1.TabIndex = 17;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.BackColor = System.Drawing.Color.Transparent;
			this.groupBox3.Controls.Add(this.cmbServicio);
			this.groupBox3.Controls.Add(this.ultraGrid2);
			this.groupBox3.Location = new System.Drawing.Point(8, 32);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(688, 152);
			this.groupBox3.TabIndex = 129;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Ingreso Servicios";
			// 
			// cmbServicio
			// 
			this.cmbServicio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbServicio.DataSource = this.cdvServicio;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 4;
			ultraGridColumn30.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbServicio.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbServicio.DisplayMember = "Nombre";
			this.cmbServicio.Location = new System.Drawing.Point(120, 40);
			this.cmbServicio.Name = "cmbServicio";
			this.cmbServicio.Size = new System.Drawing.Size(232, 56);
			this.cmbServicio.TabIndex = 123;
			this.cmbServicio.ValueMember = "idServicios";
			this.cmbServicio.Visible = false;
			// 
			// cdvServicio
			// 
			this.cdvServicio.BindingContextCtrl = this;
			this.cdvServicio.DataSet = this.cdsServicio;
			this.cdvServicio.TableName = "";
			this.cdvServicio.TableViewName = "ContratosServicios";
			// 
			// cdsServicio
			// 
			this.cdsServicio.BindingContextCtrl = this;
			this.cdsServicio.DataLibrary = "LibFacturacion";
			this.cdsServicio.DataLibraryUrl = "";
			this.cdsServicio.DataSetDef = "dsContrato";
			this.cdsServicio.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsServicio.Name = "cdsServicio";
			this.cdsServicio.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsServicio.SchemaDef = null;
			this.cdsServicio.BeforeFill += new C1.Data.FillEventHandler(this.cdsServicio_BeforeFill);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "ContratosFinanciacion";
			this.ultraGrid2.DataSource = this.cdsFinancia;
			ultraGridBand5.AddButtonCaption = "Pago";
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.Caption = "Servicio";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Width = 74;
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Width = 103;
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Width = 103;
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Width = 89;
			ultraGridColumn37.Header.Caption = "Fecha Inicio";
			ultraGridColumn37.Header.VisiblePosition = 6;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn37.Width = 114;
			ultraGridColumn38.Header.Caption = "Fecha Final";
			ultraGridColumn38.Header.VisiblePosition = 7;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn38.Width = 116;
			ultraGridColumn39.Header.VisiblePosition = 8;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 9;
			ultraGridColumn41.Header.VisiblePosition = 10;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn41});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(0, 16);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(688, 136);
			this.ultraGrid2.TabIndex = 122;
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			this.ultraGrid2.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid2_BeforeRowsDeleted);
			this.ultraGrid2.AfterCellActivate += new System.EventHandler(this.ultraGrid2_AfterCellActivate);
			// 
			// cdsFinancia
			// 
			this.cdsFinancia.BindingContextCtrl = this;
			this.cdsFinancia.DataLibrary = "LibFacturacion";
			this.cdsFinancia.DataLibraryUrl = "";
			this.cdsFinancia.DataSetDef = "dsContrato";
			this.cdsFinancia.EnforceConstraints = false;
			this.cdsFinancia.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsFinancia.Name = "cdsFinancia";
			this.cdsFinancia.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFinancia.SchemaDef = null;
			this.cdsFinancia.BeforeFill += new C1.Data.FillEventHandler(this.cdsFinancia_BeforeFill);
			// 
			// btnFactura
			// 
			this.btnFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnFactura.BackColor = System.Drawing.Color.Transparent;
			this.btnFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnFactura.Location = new System.Drawing.Point(568, 424);
			this.btnFactura.Name = "btnFactura";
			this.btnFactura.Size = new System.Drawing.Size(96, 24);
			this.btnFactura.TabIndex = 125;
			this.btnFactura.Text = "Borra Facturado";
			this.btnFactura.Click += new System.EventHandler(this.btnFactura_Click);
			// 
			// btnServicios
			// 
			this.btnServicios.BackColor = System.Drawing.Color.Transparent;
			this.btnServicios.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnServicios.Location = new System.Drawing.Point(8, 8);
			this.btnServicios.Name = "btnServicios";
			this.btnServicios.Size = new System.Drawing.Size(96, 23);
			this.btnServicios.TabIndex = 124;
			this.btnServicios.Text = "Crea Servicio";
			this.btnServicios.Click += new System.EventHandler(this.btnServicios_Click);
			// 
			// btnTabla
			// 
			this.btnTabla.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTabla.BackColor = System.Drawing.Color.Transparent;
			this.btnTabla.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnTabla.Location = new System.Drawing.Point(600, 192);
			this.btnTabla.Name = "btnTabla";
			this.btnTabla.Size = new System.Drawing.Size(88, 24);
			this.btnTabla.TabIndex = 120;
			this.btnTabla.Text = "Generar &Tabla";
			this.btnTabla.Click += new System.EventHandler(this.btnTabla_Click);
			// 
			// btnVerAdemdums
			// 
			this.btnVerAdemdums.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVerAdemdums.BackColor = System.Drawing.Color.Transparent;
			this.btnVerAdemdums.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnVerAdemdums.Location = new System.Drawing.Point(352, 496);
			this.btnVerAdemdums.Name = "btnVerAdemdums";
			this.btnVerAdemdums.Size = new System.Drawing.Size(96, 23);
			this.btnVerAdemdums.TabIndex = 143;
			this.btnVerAdemdums.Text = "Ver Ademdums";
			this.btnVerAdemdums.Visible = false;
			this.btnVerAdemdums.Click += new System.EventHandler(this.btnVerAdemdums_Click);
			// 
			// cdvMoneda
			// 
			this.cdvMoneda.BindingContextCtrl = this;
			this.cdvMoneda.DataSet = this.cdsMoneda;
			this.cdvMoneda.TableName = "";
			this.cdvMoneda.TableViewName = "MonedaTipo";
			// 
			// cdsMoneda
			// 
			this.cdsMoneda.BindingContextCtrl = this;
			this.cdsMoneda.DataLibrary = "LibContabilidad";
			this.cdsMoneda.DataLibraryUrl = "";
			this.cdsMoneda.DataSetDef = "dsMoneda";
			this.cdsMoneda.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMoneda.Name = "cdsMoneda";
			this.cdsMoneda.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsMoneda.SchemaDef = null;
			this.cdsMoneda.BeforeFill += new C1.Data.FillEventHandler(this.cdsMoneda_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Contratos";
			this.barraDato1.DataNombreId = "idContratos";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsContratos;
			this.barraDato1.DataTabla = "Contratos";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 496);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 13;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdvActividad
			// 
			this.cdvActividad.BindingContextCtrl = this;
			this.cdvActividad.DataSet = this.cdsActividad;
			this.cdvActividad.TableName = "";
			this.cdvActividad.TableViewName = "ContratosActividad";
			// 
			// cdsActividad
			// 
			this.cdsActividad.BindingContextCtrl = this;
			this.cdsActividad.DataLibrary = "LibFacturacion";
			this.cdsActividad.DataLibraryUrl = "";
			this.cdsActividad.DataSetDef = "dsContrato";
			this.cdsActividad.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsActividad.Name = "cdsActividad";
			this.cdsActividad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsActividad.SchemaDef = null;
			// 
			// cdsAmortiza
			// 
			this.cdsAmortiza.BindingContextCtrl = this;
			this.cdsAmortiza.DataLibrary = "LibFacturacion";
			this.cdsAmortiza.DataLibraryUrl = "";
			this.cdsAmortiza.DataSetDef = "dsContrato";
			this.cdsAmortiza.EnforceConstraints = false;
			this.cdsAmortiza.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsAmortiza.Name = "cdsAmortiza";
			this.cdsAmortiza.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsAmortiza.SchemaDef = null;
			this.cdsAmortiza.BeforeFill += new C1.Data.FillEventHandler(this.cdsAmortiza_BeforeFill);
			// 
			// cdvContratos
			// 
			this.cdvContratos.BindingContextCtrl = this;
			this.cdvContratos.DataSet = this.cdsContratos;
			this.cdvContratos.Sort = "idContratos";
			this.cdvContratos.TableName = "";
			this.cdvContratos.TableViewName = "Contratos";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.Location = new System.Drawing.Point(80, 16);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(88, 21);
			this.ultraTextEditor2.TabIndex = 99;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Location = new System.Drawing.Point(240, 16);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(48, 21);
			this.ultraTextEditor1.TabIndex = 98;
			// 
			// txtReporteCompara
			// 
			this.txtReporteCompara.Location = new System.Drawing.Point(80, 40);
			this.txtReporteCompara.Name = "txtReporteCompara";
			this.txtReporteCompara.Size = new System.Drawing.Size(144, 21);
			this.txtReporteCompara.TabIndex = 97;
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.Location = new System.Drawing.Point(16, 72);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(56, 16);
			this.ultraLabel11.TabIndex = 8;
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.Location = new System.Drawing.Point(176, 24);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(64, 16);
			this.ultraLabel10.TabIndex = 7;
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.Location = new System.Drawing.Point(16, 48);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel8.TabIndex = 6;
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.Location = new System.Drawing.Point(16, 24);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(48, 16);
			this.ultraLabel9.TabIndex = 5;
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(232, 64);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(24, 24);
			this.btSalir.TabIndex = 21;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(80, 64);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(144, 21);
			this.cmbProyecto.TabIndex = 96;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// ultraTextEditor5
			// 
			this.ultraTextEditor5.Location = new System.Drawing.Point(88, 48);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(208, 21);
			this.ultraTextEditor5.TabIndex = 100;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.Location = new System.Drawing.Point(16, 56);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel7.TabIndex = 6;
			// 
			// ultraTextEditor4
			// 
			this.ultraTextEditor4.Location = new System.Drawing.Point(88, 16);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(48, 21);
			this.ultraTextEditor4.TabIndex = 99;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.Location = new System.Drawing.Point(16, 24);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel6.TabIndex = 5;
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(0, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(704, 480);
			this.ultraTabControl1.TabIndex = 0;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Datos Generales";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Financiación Servicios";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			this.ultraTabControl1.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.ultraTabControl1_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(700, 454);
			// 
			// cdvFinancia
			// 
			this.cdvFinancia.BindingContextCtrl = this;
			this.cdvFinancia.DataSet = this.cdsFinancia;
			this.cdvFinancia.Sort = "idFinanciacionS";
			this.cdvFinancia.TableName = "";
			this.cdvFinancia.TableViewName = "ContratosFinanciacion";
			// 
			// cdvAmortiza
			// 
			this.cdvAmortiza.BindingContextCtrl = this;
			this.cdvAmortiza.DataSet = this.cdsAmortiza;
			this.cdvAmortiza.Sort = "idAmortizacionC";
			this.cdvAmortiza.TableName = "";
			this.cdvAmortiza.TableViewName = "ContratosAmortizacion";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// IngresoContratos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(704, 525);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btnVerAdemdums);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "IngresoContratos";
			this.Text = "Ingreso de Contratos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.IngresoContratos_Closing);
			this.Load += new System.EventHandler(this.IngresoContratos_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCComercial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.groupBox4.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaI)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtContratante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContratado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNum)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTriggerC)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFinquito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdemdum)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			this.groupBox7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFinancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAmortiza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporteCompara)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdvFinancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAmortiza)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void IngresoContratos_Load(object sender, System.EventArgs e)
		{									
			ultraTabControl1.Tabs[1].Visible= false;
			//			ModificaGrilla(false);
			//			txtTriggerC.Value = "0";
			if (this.vIdContrato > 0)
			{				
				this.barraDato1.IdRegistro = vIdContrato;
				this.barraDato1.ProximoId(5);
				//				this.barraDato1.EditarRegistro();
				//				this.barraDato1.Enabled=true;
				//				this.barraDato1.VisibleNuevo=false;	
				//				visible(true);
				//				ModificaGrilla(true);							
			}
			else
			{			
				//Filtro de ademdums
				string stFiltro = " And (Ademdum <> 1)";
				this.barraDato1.stFiltroExt = stFiltro;
				this.barraDato1.PosUltima();																				
			}
			if(Valarma == 1)
			{
				string stAlarma = "select alarma from ContratosAlarma";
				bool vAlarma = Funcion.bEjecutaSQL(cdsContratos, stAlarma);
				if(vAlarma == true) alarma();
			}	
		
		}
					
		private void btnActividad_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosActividad	ingAct = new IngresoContratosActividad())
			{
				ingAct.ShowDialog();
			}
			this.cdsActividad.Clear();
			this.cdsActividad.Fill();	
		}

		private void cdsMoneda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			this.cdsMoneda.Schema.Connections[0].ConnectionString = 
			//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsContratos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsContratos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnServicios_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosServicios ingCS = new IngresoContratosServicios())
			{
				ingCS.ShowDialog();
			}
			this.cdsServicio.Clear();
			this.cdsServicio.Fill();	
		}

		private void cdsCc_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{			
			//			Funcion miFuncion = new Funcion();
			//			this.cdsCc.Schema.Connections[0].ConnectionString = 
			//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{				
			try
			{
				string idCont = idContratos.Value.ToString();
				int idCuenta = 0;
				string stCuenta = "select count(*)from ContratosFinanciacion where idContrato = "+ idCont;
				idCuenta = Funcion.iEscalarSQL(cdsFinancia, stCuenta, true);		
				if(idCuenta > 0)
				{
					if (DialogResult.No == MessageBox.Show("Contrato Tiene Servicios Ingresados \n Eliminara Permanetemente", 
						"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;		
					string stFinan = "select  idFinanciacionS from ContratosFinanciacion where idContrato = "+ idCont;
					int stIdfinan = Funcion.iEscalarSQL(cdsFinancia, stFinan, true);
					string stBorraAmortiz = "delete ContratosAmortizacion where idFinanciacionS = "+ stIdfinan;			
					Funcion.EjecutaSQL(cdsFinancia, stBorraAmortiz, true);
					string stBorraFinan = "delete ContratosFinanciacion where idContrato = "+ idCont;			
					Funcion.EjecutaSQL(cdsFinancia, stBorraFinan, true);
				}
				else
				{
					if (DialogResult.No == MessageBox.Show("Eliminara Permanetemente", 
						"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;	
				}						
				this.barraDato1.BorraRegistro();
			}
			catch(Exception ex){MessageBox.Show(ex.Message, "Registro Borrados");}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (IngresoContratosBuscar	miBusqueda = new IngresoContratosBuscar())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					try
					{						
						barraDato1.IdRegistro=miBusqueda.idBusca;
						barraDato1.ProximoId(5);						
					}
					catch{}
				}
			}
		}
		private void alarma()
		{
			using (ContratosAlarma	miAlarma = new ContratosAlarma(true))
			{
				miAlarma.Size = new System.Drawing.Size(630,430);
				if (miAlarma.ShowDialog() == DialogResult.OK)
				{
					try
					{						
						barraDato1.IdRegistro = miAlarma.idAlarma;
						barraDato1.ProximoId(5);						
					}
					catch{}
				}
			}
		}
		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			visible(true);
			ModificaGrilla(true);
			drContrato = cdsContratos.TableViews["Contratos"].Rows[0];
			drContrato["FechaP"]=DateTime.Today;
			
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.errorProvider.SetError(this.btFinanciacion, "");								
			this.barraDato1.DeshacerRegistro();
			visible(false);
			ModificaGrilla(false);
			
//			this.barraDato1.GrabaRegistro();
//			if(vIdContrato > 0)
//			{		
//				this.DialogResult = DialogResult.OK;
//			}

		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
			visible(true);
			ModificaGrilla(true);

		}
		
		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			drContrato = cdsContratos.TableViews["Contratos"].Rows[0];				
			drContrato["Ademdum"]=0;
			if(vIdContrato>0)
			{				
				drContrato["Ademdum"]=1;
			}
			//grabaGrilla2();
			
			this.barraDato1.GrabaRegistro();
			if(vIdContrato > 0)
			{		
				this.DialogResult = DialogResult.OK;
			}
			visible(false);
					ModificaGrilla(false);
		}
		
		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{

				int VrevisaAdemdums = revisaAdemdums();			
				if(VrevisaAdemdums > 0 )
				{
					btnVerAdemdums.Visible=true;
				}
				else{btnVerAdemdums.Visible=false;}


				////grilla de ContratoFinanciacion
				drContrato = cdsContratos.TableViews["Contratos"].Rows[0];			
				C1.Data.FilterCondition fcFinanciacion = new 
					C1.Data.FilterCondition(cdsFinancia, "ContratosFinanciacion", "idContrato = " + drContrato["idContratos"]);
				/////grilla Amortizacion
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcFinanciacion);
				cdsFinancia.Fill(fcTotal, false);
			}
			catch{};

		}

		public void visible(bool ver)
		{
			groupBox1.Enabled = ver;
			groupBox2.Enabled = ver;
			groupBox4.Enabled = ver;
			txtNum.Enabled=ver;
			cmbCComercial.Enabled=ver;
			cmbCliente.Enabled = ver;
			cmbLocal.Enabled = ver;
			cmbFechaP.Enabled = ver;		
		}		
		private void calculaFecha()
		{			
			if (!barraDato1.bEditar)return;			
			if(txtAnio.Value == System.DBNull.Value)return;
			if(txtMes.Value == System.DBNull.Value)return;
			if(cmbFechaI.Value == System.DBNull.Value)return;						
			DateTime dtFecha = (DateTime) cmbFechaI.Value;
			DateTime dtFechaF; 
			int Vmes = 0;			
			int Vanio = 0;
			Vanio= int.Parse(txtAnio.Value.ToString());							
			Vmes = int.Parse(txtMes.Value.ToString());
			drContrato = cdsContratos.TableViews["Contratos"].Rows[0];
			dtFechaF = dtFecha.AddMonths(Vmes).AddYears(Vanio);
			drContrato["anio"]=Vanio;
			drContrato["mes"]=Vmes;
			drContrato["FechaF"]=dtFechaF;						
		}

		private void ultraTabPageControl2_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void btFinanciacion_Click(object sender, System.EventArgs e)
		{
			int id=0;
			try
			{
				this.errorProvider.SetError(this.btFinanciacion, "");								
				id = int.Parse(idContratos.Text);
				if(id < 1)
				{
					this.errorProvider.SetError(this.btFinanciacion, "Debe Grabar el contrato");
					return;											
				}				
			}
			catch{}

			using (IngresoContratosFinanciacion	ingFin = new IngresoContratosFinanciacion(id))
			{
				if (ingFin.ShowDialog() == DialogResult.OK)
				{
					try
					{						
						//MessageBox.Show("Generado...");
					}
					catch{}
				}
				
			}
			//this.cdsActividad.Clear();
			//this.cdsActividad.Fill();	
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFechaF_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if(e.KeyChar == 13)
			{
				int ret = calculaFecha((DateTime)cmbFechaI.Value, (DateTime)cmbFechaF.Value,1);
			}
		}

		private void chkCliente_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int stIdContratos =int.Parse(idContratos.Text);
			if (vIdContrato>0)
			{
				int idC;
				string stSelect = "select idPadre from contratos where idContratos = "+vIdContrato;
				idC = Funcion.iEscalarSQL(cdsFinancia,stSelect,true);
				stIdContratos = idC;
			}
			
			if (idContratos.Text.Length > 0 )
			{
				int verif = 0;
				string stSelect = "select count(*) from contratos where idPadre = "+stIdContratos +" or idContratos ="+stIdContratos ;
				verif = Funcion.iEscalarSQL(cdsFinancia,stSelect,true);

				if(chkCliente.Checked == true)
				{				
			
					if (verif >= 1)
					{
						string stUpdate = "update Contratos set Finiquito = 1, FechaFiniquito = getdate() where idPadre = "+ stIdContratos  +" or idContratos ="+stIdContratos ;;			
						Funcion.EjecutaSQL(cdsFinancia, stUpdate, true);			
					}				
				}
				else
				{
					
					if (verif >= 1)
					{
						string stUpdate = "update Contratos set Finiquito = 0 where idPadre = "+ stIdContratos +" or idContratos ="+stIdContratos ;;			
						Funcion.EjecutaSQL(cdsFinancia, stUpdate, true);														
					}
				}
			}
		}


		private void cdsFinancia_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			this.cdsFinancia.Schema.Connections[0].ConnectionString = 
			//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsAmortiza_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAmortiza.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		C1.Data.C1DataRow drContrato;
		C1.Data.C1DataRow drContratoA;

		private void cdsServicio_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			this.cdsServicio.Schema.Connections[0].ConnectionString = 
			//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		

		private void control1()
		{
			try
			{
				string va;
				va = cmbCComercial.Value.ToString();
				if (va.Length <= 0)
				{					
					return;
				}
				else
				{			
					cdvLocales.RowFilter="idCentroComercial = "+ va;
				}			
			}
			catch{}
		}

		private void txtAnio_ValueChanged(object sender, System.EventArgs e)
		{
			calculaFecha();
		}

		private void txtMes_ValueChanged(object sender, System.EventArgs e)
		{
			calculaFecha();
		}		
		private void ultraGrid2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			drContrato = cdsContratos.TableViews["Contratos"].Rows[0];
			e.Row.Cells["idContrato"].Value = drContrato["idContratos"];
		}	
		private void btnTabla_Click(object sender, System.EventArgs e)
		{
			if (vIdContrato <= 0)
			{
				txtTriggerC.Value = 0;
			}

			if(ultraGrid2.Rows.Count < 1 )return;
			Cursor = Cursors.WaitCursor;
			try
			{						
				int Id = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;	
				int idSericio = (int) ultraGrid2.ActiveRow.Cells["idServicio"].Value;
//				string vTriger = txtTriggerC.Value.ToString();
				if(Id < 0)return;
				generaAmortizacion(Id,1,idSericio,"0");
			}
			catch(Exception ex){MessageBox.Show("Error: "+ex);}
			Cursor = Cursors.Default;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			drContratoA = cdsFinancia.TableViews["ContratosFinanciacion"].Rows[0];
			e.Row.Cells["idFinanciacionS"].Value = drContratoA["idFinanciacionS"];
		}		
		private void ultraGrid2_AfterCellActivate(object sender, System.EventArgs e)
		{
			
			int Id = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;
			if(Id < 0)return;
			

			generaAmortizacion(Id,2,0,"0");
		}	
		
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		
		private void generaAmortizacion(int idFinancia, int controlP, int idServicio, string vTriger)
		{			
			int verif = verificaDatosAmortiza(idFinancia);			
			if (verif > 0 && controlP==1 )return;
			//			if(controlP == 1)
			//			{
			//				if (DialogResult.No == MessageBox.Show("Desea Generar Tabla de Amortizacion", 
			//					"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			//			}
			
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ContratosAmortiza "+idFinancia+" ,"+controlP+" ,"+idContratos.Text+" ,"+idServicio+" ,"+vTriger;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;		
			
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idAmortizacionC"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idFinanciacionS"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["FechaCobro"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["CapCobrado"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["Facturado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			

			#endregion Visibilidad
		}

		private int verificaDatosAmortiza(int idFinancia)
		{	
			int idCuenta = 0;	
			if(vIdContrato > 0 )
			{
				return idCuenta;
			}
			else
			{			
				string stCuenta = "select count(*)from ContratosAmortizacion where idFinanciacionS = "+ idFinancia;
				idCuenta = Funcion.iEscalarSQL(cdsFinancia, stCuenta, true);
				return idCuenta;
			}
		}

		private void ultraTabControl1_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
			try
			{
				this.errorProvider.SetError(this.barraDato1, "");								
				int id = int.Parse(idContratos.Text);
				if(id < 1)
				{
					this.errorProvider.SetError(this.barraDato1, "Debe Grabar el contrato");								
					ultraTabControl1.SelectedTab = ultraTabControl1.Tabs[0];				
				}
				if(ultraTabControl1.Tabs[1].Selected == true)
				{ 
					//barraDato1.VisibleNuevo=false;
				}
			}
			catch{}


		}

		private void btnFactura_Click(object sender, System.EventArgs e)
		{			
			int idAmort = (int) ultraGrid1.ActiveRow.Cells["idAmortizacionC"].Value;
			int idfin = (int) ultraGrid1.ActiveRow.Cells["idFinanciacionS"].Value;
			string stUpdate = "update ContratosAmortizacion set facturado = 0 where idAmortizacionC = "+ idAmort;			
			Funcion.EjecutaSQL(cdsFinancia, stUpdate, true);			
			generaAmortizacion(idfin,2,0,"0");
		}		

		private int verificaFilas()
		{
			int cont=1;	
			bool verifica;		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{		//cont=1;			
				try 
				{
					verifica = (bool)dr.Cells["Facturar"].Value;
					if(verifica == true)
					{		cont = cont+cont;}
				}
				catch{}				
			}
			return cont;
		}

		private void ultraGrid2_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg=false;
			if (DialogResult.No == MessageBox.Show("Desea Eliminar Registro", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			try
			{
				this.errorProvider.SetError(this.ultraGrid2, "");								
				
				int idFinac = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;
				int id= verificaDatosAmortiza(idFinac);
				if(id > 0)
				{
					MessageBox.Show("Registro Tiene Amortizacion Generada","No puede Eliminar",MessageBoxButtons.OK,MessageBoxIcon.Stop);
					e.Cancel=true;
					return;
				}
				int idFin = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;
				string stCuenta = "delete ContratosFinanciacion where idFinanciacionS = "+ idFin;			
				Funcion.EjecutaSQL(cdsFinancia, stCuenta, true);
			}
			catch(Exception ex){MessageBox.Show("Error al borrar: "+ex);}
			return;
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg=false;
			if (DialogResult.No == MessageBox.Show("Desea Eliminar Registro", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			borraAmortizacion();
		}
		private void borraAmortizacion()
		{
			int idAmort = (int) ultraGrid1.ActiveRow.Cells["idFinanciacionS"].Value;
			string stCuenta = "delete ContratosAmortizacion where idFinanciacionS = "+ idAmort;			
			Funcion.EjecutaSQL(cdsFinancia, stCuenta, true);
		}

		private int revisaAdemdums()
		{			
			int valor = 0;
			try
			{
				string stRevisa = "exec ContratosverificaAdemdum "+ idContratos.Text;			
				valor = Funcion.iEscalarSQL(cdsFinancia, stRevisa, true);	
			}
			catch{}
		return valor;
		}
		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{			

			int VrevisaAdemdums = revisaAdemdums();
			string VidContratos;
			if(VrevisaAdemdums <=0 )
			{
				VidContratos = this.idContratos.Value.ToString();
			}
			else
			{
				VidContratos = VrevisaAdemdums.ToString();
			}
				this.Cursor = Cursors.WaitCursor;
				string stFiltro = "{Contratos.idContratos} = " + VidContratos;
				string stReporte = "Contratos.rpt";			
				Reporte miReporte = new Reporte(stReporte, stFiltro);
			
				miReporte.MdiParent = this.MdiParent;
				miReporte.Show();
			
				this.Cursor = Cursors.Default;
			
		}

		private void btnLocal_Click_1(object sender, System.EventArgs e)
		{			
			using (IngresoContratoLocales	inglocal = new IngresoContratoLocales())
			{
				if (inglocal.ShowDialog() == DialogResult.OK)
				{
					try
					{			
						this.cdsCc.Clear();
						this.cdsCc.Fill();
						this.cdsLocales.Clear();
						this.cdsLocales.Fill();		
					}
					catch{}
				}
			}		
		}

		private void cmbCComercial_ValueChanged(object sender, System.EventArgs e)
		{			
			control1();		
		}

		private void IngresoContratos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				if (DialogResult.No == MessageBox.Show("No se Graban los Cambios... \n Desea Salir sin Grabar", "Advertencia",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					e.Cancel = true;
					return;
				}				
			}
		}

	
		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid2.Enabled = true;
			if (bModifica)
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		
		private void grabaGrilla2()
		{
//			int verif = 0;			
//			int Id =0;	
//			if(vIdContrato > 0)
//			{		
//				//        copia la tabla contratos org a una Aux	
//				if(ultraGrid2.ActiveRow != null)
//					Id = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;	
//				if(Id < 0)return;		
//				//				string stCuenta = "exec ContratoCopiar "+ vIdContrato;
//				//				Funcion.EjecutaSQL(cdsContratos, stCuenta, true);
////				this.DialogResult = DialogResult.OK;
//				txtTriggerC.Value = "1";
//				verif = 1;				
//			}			                       
//			this.barraDato1.GrabaRegistro();			
			

			if(ultraGrid2.Rows.Count<1)return;			
			try
			{
				int cuotas = 0;
				Cursor = Cursors.WaitCursor;
				if(ultraGrid2.ActiveRow.Cells["Cuotas"].Value != System.DBNull.Value)
				{
					cuotas = (int) ultraGrid2.ActiveRow.Cells["Cuotas"].Value;
				}
//				else 
//				{
//					
//				  ultraGrid2.ActiveRow.Cells["VTrigger"].Value = cuotas;
//				}				
				if(ultraGrid2.ActiveRow.Cells["FechaD"].Value == System.DBNull.Value)
				{
					ultraGrid2.ActiveRow.Cells["FechaD"].Value = DateTime.Today;				
				}
				
				if ((ultraGrid2.ActiveRow.Cells["FechaH"].Value == System.DBNull.Value))
				{
					ultraGrid2.ActiveRow.Cells["FechaH"].Value = DateTime.Today.AddMonths(cuotas);
						ultraGrid2.ActiveRow.Cells["Cuotas"].Value = cuotas;
				}else
				{
					int ret1 = calculaFecha((DateTime)ultraGrid2.ActiveRow.Cells["FechaD"].Value, (DateTime)ultraGrid2.ActiveRow.Cells["FechaH"].Value, 2);
					
					if(ret1 > 0)
					{
						ultraGrid2.ActiveRow.Cells["Cuotas"].Value = ret1;
					}
				
				}
								
				//ultraGrid2.ActiveRow.Cells["VTrigger"].Value = verif;				
			}
			catch (Exception ex)
			{	MessageBox.Show("Error "+ex);	}		
//			ModificaGrilla(false);
//			if(vIdContrato > 0) 
//			{
//				generaAmortizacion(Id,1);
//			}	
			this.cdsFinancia.Update();	
			
			Cursor = Cursors.Default;
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			using (Cliente cli = new Cliente(true))
			{
				cli.ShowDialog();
			}
			this.cdsCliente.Clear();
			this.cdsCliente.Fill();		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			grabaGrilla2();
		
		}

		private void btnVerAdemdums_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosAdemdums	miBusquedaA = new IngresoContratosAdemdums(int.Parse(idContratos.Text)))
			{
				if (miBusquedaA.ShowDialog() == DialogResult.OK)
				{
					try
					{								
						IngresoContratos contratosB = new IngresoContratos(miBusquedaA.idBusca);
						contratosB.Show();
					}
					catch{}
				}
			}
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea crear Ademdum?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			Cursor = Cursors.WaitCursor;
string cc="";
			using (IngresoContratoAdemdumdet	miBusquedaAd = new IngresoContratoAdemdumdet(int.Parse(idContratos.Text)))
			{
				if (miBusquedaAd.ShowDialog() == DialogResult.OK)
				{
					try
					{								
						cc=miBusquedaAd.texto;
					}
					catch{}
				}
			}


			string stCopiar = "exec ContratosCopiar "+ idContratos.Text+", '"+cc+"'";		
			int id = Funcion.iEscalarSQL(cdsFinancia, stCopiar, true);					
			Cursor = Cursors.Default;
			IngresoContratos contratosCopiar = new IngresoContratos(id);
			contratosCopiar.Show();
			
		}

		private void cmbFechaF_TextChanged(object sender, System.EventArgs e)
		{

//		   int ret = calculaFecha((DateTime)cmbFechaI.Value, (DateTime)cmbFechaF.Value,1);
		}		
		private int calculaFecha(DateTime desde, DateTime hasta, int Control)
		{
			int[] monthDay = new int[12] { 31, -1, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
			int year;
			int month;
			int day;
			DateTime dtD; DateTime dtH;
			dtD = desde;
			dtH = hasta;
           
			int increment;

			if (dtD > dtH)
			{
				MessageBox.Show("Fecha Hasta debe ser Mayor", "Fechas", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return 0;
			}

			/// 
			/// Day Calculation
			/// 
			increment = 0;

			if (dtD.Day > dtH.Day)
			{
				increment = monthDay[dtD.Month - 1];

			}
			if (increment == -1)
			{
				if (DateTime.IsLeapYear(dtD.Year))
				{
					// leap year february contain 29 days
					increment = 29;
				}
				else
				{
					increment = 28;
				}
			}
			if (increment != 0)
			{
				day = (dtH.Day + increment) - dtD.Day;
				increment = 1;
			}
			else
			{
				day = dtH.Day - dtD.Day;
			}

			///
			///month calculation
			///
			if ((dtD.Month + increment) > dtH.Month)
			{
				month = (dtH.Month + 12) - (dtD.Month + increment);
				increment = 1;
			}
			else
			{
				month = (dtH.Month) - (dtD.Month + increment);
				increment = 0;
			}

			///
			/// year calculation
			///
			year = dtH.Year - (dtD.Year + increment);
			//MessageBox.Show(" anio " + year + " mes " + month + " dia " + day);
			int retorna = 0;
			if(Control == 1)
			{
//				txtAnio.Value =  year;
//				txtMes.Value = month;			
				drContrato = cdsContratos.TableViews["Contratos"].Rows[0];				
				drContrato["anio"]=year;
				drContrato["mes"]=month;
		
			}
			else
			{
				retorna = ((year*12)+month);
			}
			return retorna; 
		}	
	}
}
