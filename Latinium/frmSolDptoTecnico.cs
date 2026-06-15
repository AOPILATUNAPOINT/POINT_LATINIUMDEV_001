using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmSolDptoTecnico.
	/// </summary>
	public class frmSolDptoTecnico : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgTipoConsulta;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTaller;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor7;
		private System.Windows.Forms.Label lblSubTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRevisado;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpcConGarantia;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpcSinGarantia;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcGarantia;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox5;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtRetiro;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEnvio;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInforme;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEntrega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtGestionCliente;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.Misc.UltraButton btNotaCrédito;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbContacto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbContactoCliente;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private Infragistics.Win.Misc.UltraButton btNuevo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObsIngreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDiagnostico;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTecnicoSinGarantia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNroOrden;
		private System.Windows.Forms.GroupBox gbCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.GroupBox gbProducto;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSerial;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumFactura;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdIngreso;
		private Infragistics.Win.Misc.UltraButton btGuardar;
		private System.Windows.Forms.GroupBox gbProveedor;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = false;
		private Infragistics.Win.Misc.UltraButton btEditar;
		private Infragistics.Win.Misc.UltraButton btConsultar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesIDT;
		private C1.Data.C1DataSet cdsDTContacto;
		private System.Windows.Forms.Label lblNumeroEgreso;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.ContextMenu mnuTipoInforme;
		private System.Windows.Forms.MenuItem mnuTipoInformeIngreso;
		private System.Windows.Forms.MenuItem mnuTipoInformeEgreso;
		private DevExpress.XtraEditors.RadioGroup rdgOrigen;
		private DevExpress.XtraEditors.ComboBoxEdit cmbGarantia;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.Misc.UltraButton btIngresoDT;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInforme;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDCFacVenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtGrupo;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIngreso;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpcGestionesTalleres;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpcGestionesClientes;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAccesorios;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.Label label35;
		private DevExpress.XtraEditors.ComboBoxEdit cmbEstados;
		private DevExpress.XtraEditors.RadioGroup radioGroup1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		bool bEdicion = false;

		public frmSolDptoTecnico()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestionCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGestionCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("DTContactoCliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContactoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestionCliente", -1, "cmbContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSolDptoTecnico));
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicioTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idServicioTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestionCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("DTContactoCliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContactoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestionCliente", -1, "cmbContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton7 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton8 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			this.utpcConGarantia = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.radioGroup1 = new DevExpress.XtraEditors.RadioGroup();
			this.label35 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbContacto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsDTContacto = new C1.Data.C1DataSet();
			this.label31 = new System.Windows.Forms.Label();
			this.dtGestionCliente = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label27 = new System.Windows.Forms.Label();
			this.txtObservacionesIDT = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label24 = new System.Windows.Forms.Label();
			this.txtInforme = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.dtRetiro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label16 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbContactoCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtEnvio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label22 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.dtInforme = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNroOrden = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTaller = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbRevisado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbEstados = new DevExpress.XtraEditors.ComboBoxEdit();
			this.utpcSinGarantia = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.ultraTextEditor7 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTecnicoSinGarantia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label28 = new System.Windows.Forms.Label();
			this.utpcGestionesTalleres = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.utpcGestionesClientes = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btNotaCrédito = new Infragistics.Win.Misc.UltraButton();
			this.btIngresoDT = new Infragistics.Win.Misc.UltraButton();
			this.rdgTipoConsulta = new DevExpress.XtraEditors.RadioGroup();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtObsIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.txtDiagnostico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.utcGarantia = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.btNuevo = new Infragistics.Win.Misc.UltraButton();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.gbCliente = new System.Windows.Forms.GroupBox();
			this.label39 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.gbProducto = new System.Windows.Forms.GroupBox();
			this.txtNumIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.txtidDCFacVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdSerial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCodArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbGarantia = new DevExpress.XtraEditors.ComboBoxEdit();
			this.gbProveedor = new System.Windows.Forms.GroupBox();
			this.txtGrupo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbFechaFactura = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.btGuardar = new Infragistics.Win.Misc.UltraButton();
			this.btEditar = new Infragistics.Win.Misc.UltraButton();
			this.btConsultar = new Infragistics.Win.Misc.UltraButton();
			this.lblNumeroEgreso = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.mnuTipoInforme = new System.Windows.Forms.ContextMenu();
			this.mnuTipoInformeIngreso = new System.Windows.Forms.MenuItem();
			this.mnuTipoInformeEgreso = new System.Windows.Forms.MenuItem();
			this.rdgOrigen = new DevExpress.XtraEditors.RadioGroup();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label34 = new System.Windows.Forms.Label();
			this.txtAccesorios = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.utpcConGarantia.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDTContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtGestionCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesIDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRetiro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbContactoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEnvio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroOrden)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTaller)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRevisado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).BeginInit();
			this.utpcSinGarantia.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTecnicoSinGarantia)).BeginInit();
			this.utpcGestionesTalleres.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObsIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiagnostico)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.utcGarantia)).BeginInit();
			this.utcGarantia.SuspendLayout();
			this.gbCliente.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.gbProducto.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDCFacVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGarantia.Properties)).BeginInit();
			this.gbProveedor.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOrigen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			this.SuspendLayout();
			// 
			// utpcConGarantia
			// 
			this.utpcConGarantia.Controls.Add(this.radioGroup1);
			this.utpcConGarantia.Controls.Add(this.label35);
			this.utpcConGarantia.Controls.Add(this.groupBox4);
			this.utpcConGarantia.Controls.Add(this.groupBox1);
			this.utpcConGarantia.Controls.Add(this.cmbContacto);
			this.utpcConGarantia.Controls.Add(this.ultraGrid1);
			this.utpcConGarantia.Controls.Add(this.label31);
			this.utpcConGarantia.Controls.Add(this.dtGestionCliente);
			this.utpcConGarantia.Controls.Add(this.label27);
			this.utpcConGarantia.Controls.Add(this.txtObservacionesIDT);
			this.utpcConGarantia.Controls.Add(this.dtIngreso);
			this.utpcConGarantia.Controls.Add(this.label24);
			this.utpcConGarantia.Controls.Add(this.txtInforme);
			this.utpcConGarantia.Controls.Add(this.label26);
			this.utpcConGarantia.Controls.Add(this.dtRetiro);
			this.utpcConGarantia.Controls.Add(this.label16);
			this.utpcConGarantia.Controls.Add(this.label25);
			this.utpcConGarantia.Controls.Add(this.label15);
			this.utpcConGarantia.Controls.Add(this.cmbContactoCliente);
			this.utpcConGarantia.Controls.Add(this.dtEnvio);
			this.utpcConGarantia.Controls.Add(this.label22);
			this.utpcConGarantia.Controls.Add(this.label18);
			this.utpcConGarantia.Controls.Add(this.label20);
			this.utpcConGarantia.Controls.Add(this.label21);
			this.utpcConGarantia.Controls.Add(this.label23);
			this.utpcConGarantia.Controls.Add(this.dtInforme);
			this.utpcConGarantia.Controls.Add(this.label17);
			this.utpcConGarantia.Controls.Add(this.txtNroOrden);
			this.utpcConGarantia.Controls.Add(this.dtEntrega);
			this.utpcConGarantia.Controls.Add(this.cmbTaller);
			this.utpcConGarantia.Controls.Add(this.ultraTextEditor6);
			this.utpcConGarantia.Controls.Add(this.cmbRevisado);
			this.utpcConGarantia.Controls.Add(this.label19);
			this.utpcConGarantia.Controls.Add(this.cmbEstados);
			this.utpcConGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.utpcConGarantia.Location = new System.Drawing.Point(1, 20);
			this.utpcConGarantia.Name = "utpcConGarantia";
			this.utpcConGarantia.Size = new System.Drawing.Size(1222, 339);
			this.utpcConGarantia.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// radioGroup1
			// 
			this.radioGroup1.EditValue = -1;
			this.radioGroup1.Location = new System.Drawing.Point(88, 272);
			this.radioGroup1.Name = "radioGroup1";
			// 
			// radioGroup1.Properties
			// 
			this.radioGroup1.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.radioGroup1.Properties.Appearance.Options.UseBackColor = true;
			this.radioGroup1.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Pendiente"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Recibido")});
			this.radioGroup1.Size = new System.Drawing.Size(168, 32);
			this.radioGroup1.TabIndex = 320;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(8, 240);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(38, 17);
			this.label35.TabIndex = 319;
			this.label35.Text = "Estado";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(0, 224);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(712, 8);
			this.groupBox4.TabIndex = 318;
			this.groupBox4.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(712, 8);
			this.groupBox1.TabIndex = 317;
			this.groupBox1.TabStop = false;
			// 
			// cmbContacto
			// 
			this.cmbContacto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbContacto.DataSource = this.ultraDataSource4;
			this.cmbContacto.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 191;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 254;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbContacto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbContacto.DisplayMember = "Descripcion";
			this.cmbContacto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbContacto.Location = new System.Drawing.Point(928, 152);
			this.cmbContacto.Name = "cmbContacto";
			this.cmbContacto.Size = new System.Drawing.Size(256, 112);
			this.cmbContacto.TabIndex = 316;
			this.cmbContacto.ValueMember = "idGestionCliente";
			this.cmbContacto.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "DTContactoCliente";
			this.ultraGrid1.DataSource = this.cdsDTContacto;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Band 0";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 89;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 49;
			ultraGridColumn5.Header.Caption = "Fecha";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.PromptChar = ' ';
			ultraGridColumn5.Width = 68;
			ultraGridColumn6.Header.Caption = "Gestión";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.PromptChar = ' ';
			ultraGridColumn6.Width = 97;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn7.Width = 239;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.PromptChar = ' ';
			ultraGridColumn8.Width = 71;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 72;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 99;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "Gestión con cliente";
			ultraGridBand2.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(720, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 320);
			this.ultraGrid1.TabIndex = 315;
			// 
			// cdsDTContacto
			// 
			this.cdsDTContacto.BindingContextCtrl = this;
			this.cdsDTContacto.DataLibrary = "LibPersonal";
			this.cdsDTContacto.DataLibraryUrl = "";
			this.cdsDTContacto.DataSetDef = "dsDTContactoCliente";
			this.cdsDTContacto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDTContacto.Name = "cdsDTContacto";
			this.cdsDTContacto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsDTContacto.SchemaDef = null;
			this.cdsDTContacto.BeforeFill += new C1.Data.FillEventHandler(this.cdsDTContacto_BeforeFill);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(808, 16);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(90, 17);
			this.label31.TabIndex = 313;
			this.label31.Text = "Fecha de Gestión";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtGestionCliente
			// 
			this.dtGestionCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtGestionCliente.DateButtons.Add(dateButton1);
			this.dtGestionCliente.Format = "dd/MM/yyyy";
			this.dtGestionCliente.Location = new System.Drawing.Point(912, 16);
			this.dtGestionCliente.Name = "dtGestionCliente";
			this.dtGestionCliente.NonAutoSizeHeight = 23;
			this.dtGestionCliente.Size = new System.Drawing.Size(104, 21);
			this.dtGestionCliente.SpinButtonsVisible = true;
			this.dtGestionCliente.TabIndex = 312;
			this.dtGestionCliente.Value = ((object)(resources.GetObject("dtGestionCliente.Value")));
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label27.Location = new System.Drawing.Point(1192, 467);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(66, 17);
			this.label27.TabIndex = 131;
			this.label27.Text = "FALTANTE:";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservacionesIDT
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesIDT.Appearance = appearance8;
			this.txtObservacionesIDT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesIDT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesIDT.Location = new System.Drawing.Point(416, 240);
			this.txtObservacionesIDT.Multiline = true;
			this.txtObservacionesIDT.Name = "txtObservacionesIDT";
			this.txtObservacionesIDT.Size = new System.Drawing.Size(288, 88);
			this.txtObservacionesIDT.TabIndex = 304;
			// 
			// dtIngreso
			// 
			this.dtIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtIngreso.DateButtons.Add(dateButton2);
			this.dtIngreso.Format = "dd/MM/yyyy";
			this.dtIngreso.Location = new System.Drawing.Point(368, 46);
			this.dtIngreso.Name = "dtIngreso";
			this.dtIngreso.NonAutoSizeHeight = 23;
			this.dtIngreso.Size = new System.Drawing.Size(104, 21);
			this.dtIngreso.SpinButtonsVisible = true;
			this.dtIngreso.TabIndex = 294;
			this.dtIngreso.Value = ((object)(resources.GetObject("dtIngreso.Value")));
			this.dtIngreso.Leave += new System.EventHandler(this.dtIngreso_Leave);
			this.dtIngreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtIngreso_KeyPress);
			this.dtIngreso.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtIngreso_BeforeDropDown);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(808, 48);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(107, 17);
			this.label24.TabIndex = 306;
			this.label24.Text = "Contacto con Cliente";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInforme
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInforme.Appearance = appearance9;
			this.txtInforme.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInforme.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInforme.Location = new System.Drawing.Point(8, 104);
			this.txtInforme.Multiline = true;
			this.txtInforme.Name = "txtInforme";
			this.txtInforme.Size = new System.Drawing.Size(696, 88);
			this.txtInforme.TabIndex = 298;
			this.txtInforme.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtInforme_KeyPress);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Enabled = false;
			this.label26.Location = new System.Drawing.Point(808, 208);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(92, 17);
			this.label26.TabIndex = 311;
			this.label26.Text = "Fecha de Entrega";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtRetiro
			// 
			this.dtRetiro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtRetiro.DateButtons.Add(dateButton3);
			this.dtRetiro.Format = "dd/MM/yyyy";
			this.dtRetiro.Location = new System.Drawing.Point(120, 200);
			this.dtRetiro.Name = "dtRetiro";
			this.dtRetiro.NonAutoSizeHeight = 23;
			this.dtRetiro.Size = new System.Drawing.Size(104, 21);
			this.dtRetiro.SpinButtonsVisible = true;
			this.dtRetiro.TabIndex = 300;
			this.dtRetiro.Value = ((object)(resources.GetObject("dtRetiro.Value")));
			this.dtRetiro.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtRetiro_KeyPress);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(328, 8);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(79, 17);
			this.label16.TabIndex = 291;
			this.label16.Text = "Fecha de envio";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(808, 88);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(76, 17);
			this.label25.TabIndex = 309;
			this.label25.Text = "Observaciones";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(8, 8);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(32, 17);
			this.label15.TabIndex = 288;
			this.label15.Text = "Taller";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbContactoCliente
			// 
			this.cmbContactoCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbContactoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbContactoCliente.DisplayLayout.AutoFitColumns = true;
			this.cmbContactoCliente.DisplayMember = "";
			this.cmbContactoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbContactoCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbContactoCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbContactoCliente.Location = new System.Drawing.Point(904, 48);
			this.cmbContactoCliente.MaxDropDownItems = 30;
			this.cmbContactoCliente.Name = "cmbContactoCliente";
			this.cmbContactoCliente.Size = new System.Drawing.Size(280, 21);
			this.cmbContactoCliente.TabIndex = 307;
			this.cmbContactoCliente.ValueMember = "";
			// 
			// dtEnvio
			// 
			this.dtEnvio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtEnvio.DateButtons.Add(dateButton4);
			this.dtEnvio.Format = "dd/MM/yyyy";
			this.dtEnvio.Location = new System.Drawing.Point(440, 8);
			this.dtEnvio.Name = "dtEnvio";
			this.dtEnvio.NonAutoSizeHeight = 23;
			this.dtEnvio.Size = new System.Drawing.Size(104, 21);
			this.dtEnvio.SpinButtonsVisible = true;
			this.dtEnvio.TabIndex = 290;
			this.dtEnvio.Value = ((object)(resources.GetObject("dtEnvio.Value")));
			this.dtEnvio.Leave += new System.EventHandler(this.dtEnvio_Leave);
			this.dtEnvio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtEnvio_KeyPress);
			this.dtEnvio.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtEnvio_BeforeDropDown);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(8, 312);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(69, 17);
			this.label22.TabIndex = 302;
			this.label22.Text = "Revisado por";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(256, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(89, 17);
			this.label18.TabIndex = 295;
			this.label18.Text = "Fecha de ingreso";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(8, 80);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(44, 17);
			this.label20.TabIndex = 299;
			this.label20.Text = "Informe";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(8, 200);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(79, 17);
			this.label21.TabIndex = 301;
			this.label21.Text = "Fecha de retiro";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(328, 240);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(76, 17);
			this.label23.TabIndex = 305;
			this.label23.Text = "Observaciones";
			// 
			// dtInforme
			// 
			this.dtInforme.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtInforme.DateButtons.Add(dateButton5);
			this.dtInforme.Format = "dd/MM/yyyy";
			this.dtInforme.Location = new System.Drawing.Point(600, 46);
			this.dtInforme.Name = "dtInforme";
			this.dtInforme.NonAutoSizeHeight = 23;
			this.dtInforme.Size = new System.Drawing.Size(104, 21);
			this.dtInforme.SpinButtonsVisible = true;
			this.dtInforme.TabIndex = 296;
			this.dtInforme.Value = ((object)(resources.GetObject("dtInforme.Value")));
			this.dtInforme.Leave += new System.EventHandler(this.dtInforme_Leave);
			this.dtInforme.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtInforme_KeyPress);
			this.dtInforme.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtInforme_BeforeDropDown);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(8, 48);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(49, 17);
			this.label17.TabIndex = 292;
			this.label17.Text = "N. Ordén";
			// 
			// txtNroOrden
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNroOrden.Appearance = appearance10;
			this.txtNroOrden.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNroOrden.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNroOrden.Location = new System.Drawing.Point(80, 45);
			this.txtNroOrden.Name = "txtNroOrden";
			this.txtNroOrden.Size = new System.Drawing.Size(160, 22);
			this.txtNroOrden.TabIndex = 293;
			this.txtNroOrden.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNroOrden_KeyPress);
			// 
			// dtEntrega
			// 
			this.dtEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtEntrega.DateButtons.Add(dateButton6);
			this.dtEntrega.Format = "dd/MM/yyyy";
			this.dtEntrega.Location = new System.Drawing.Point(912, 208);
			this.dtEntrega.Name = "dtEntrega";
			this.dtEntrega.NonAutoSizeHeight = 23;
			this.dtEntrega.Size = new System.Drawing.Size(104, 21);
			this.dtEntrega.SpinButtonsVisible = true;
			this.dtEntrega.TabIndex = 310;
			this.dtEntrega.Value = ((object)(resources.GetObject("dtEntrega.Value")));
			// 
			// cmbTaller
			// 
			this.cmbTaller.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTaller.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTaller.DataSource = this.ultraDataSource2;
			this.cmbTaller.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 122;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 237;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbTaller.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTaller.DisplayMember = "Nombre";
			this.cmbTaller.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTaller.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTaller.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTaller.Location = new System.Drawing.Point(56, 8);
			this.cmbTaller.MaxDropDownItems = 30;
			this.cmbTaller.Name = "cmbTaller";
			this.cmbTaller.Size = new System.Drawing.Size(256, 21);
			this.cmbTaller.TabIndex = 289;
			this.cmbTaller.ValueMember = "idServicioTecnico";
			this.cmbTaller.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTaller_KeyPress);
			this.cmbTaller.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTaller_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ultraTextEditor6
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor6.Appearance = appearance11;
			this.ultraTextEditor6.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor6.Location = new System.Drawing.Point(808, 112);
			this.ultraTextEditor6.Multiline = true;
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(400, 80);
			this.ultraTextEditor6.TabIndex = 308;
			// 
			// cmbRevisado
			// 
			this.cmbRevisado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRevisado.DataSource = this.ultraDataSource1;
			this.cmbRevisado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 205;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbRevisado.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRevisado.DisplayMember = "Nombre";
			this.cmbRevisado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRevisado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRevisado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRevisado.Location = new System.Drawing.Point(88, 312);
			this.cmbRevisado.MaxDropDownItems = 30;
			this.cmbRevisado.Name = "cmbRevisado";
			this.cmbRevisado.Size = new System.Drawing.Size(224, 21);
			this.cmbRevisado.TabIndex = 303;
			this.cmbRevisado.ValueMember = "idPersonal";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(488, 48);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(91, 17);
			this.label19.TabIndex = 297;
			this.label19.Text = "Fecha de informe";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstados
			// 
			this.cmbEstados.Location = new System.Drawing.Point(88, 240);
			this.cmbEstados.Name = "cmbEstados";
			// 
			// cmbEstados.Properties
			// 
			this.cmbEstados.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.cmbEstados.Properties.Appearance.Options.UseFont = true;
			this.cmbEstados.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbEstados.Properties.Items.AddRange(new object[] {
																															 "REPARADO",
																															 "SIN REPARACIÓN",
																															 "NO APLICA GARANTIA",
																															 "NUEVO",
																															 "NOTA DE CRÉDITO",
																															 ""});
			this.cmbEstados.Size = new System.Drawing.Size(224, 26);
			this.cmbEstados.TabIndex = 290;
			// 
			// utpcSinGarantia
			// 
			this.utpcSinGarantia.Controls.Add(this.lblSubTotal);
			this.utpcSinGarantia.Controls.Add(this.txtIva0);
			this.utpcSinGarantia.Controls.Add(this.label29);
			this.utpcSinGarantia.Controls.Add(this.ultraTextEditor7);
			this.utpcSinGarantia.Controls.Add(this.cmbTecnicoSinGarantia);
			this.utpcSinGarantia.Controls.Add(this.label28);
			this.utpcSinGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.utpcSinGarantia.Location = new System.Drawing.Point(-10000, -10000);
			this.utpcSinGarantia.Name = "utpcSinGarantia";
			this.utpcSinGarantia.Size = new System.Drawing.Size(1222, 339);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(56, 144);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(36, 17);
			this.lblSubTotal.TabIndex = 264;
			this.lblSubTotal.Text = "Abono";
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance12;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(120, 136);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 265;
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(368, 16);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(62, 17);
			this.label29.TabIndex = 263;
			this.label29.Text = "Diagnóstico";
			// 
			// ultraTextEditor7
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor7.Appearance = appearance13;
			this.ultraTextEditor7.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor7.Location = new System.Drawing.Point(368, 40);
			this.ultraTextEditor7.Multiline = true;
			this.ultraTextEditor7.Name = "ultraTextEditor7";
			this.ultraTextEditor7.Size = new System.Drawing.Size(432, 72);
			this.ultraTextEditor7.TabIndex = 262;
			// 
			// cmbTecnicoSinGarantia
			// 
			this.cmbTecnicoSinGarantia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTecnicoSinGarantia.DataSource = this.ultraDataSource1;
			this.cmbTecnicoSinGarantia.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 261;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbTecnicoSinGarantia.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTecnicoSinGarantia.DisplayMember = "Nombre";
			this.cmbTecnicoSinGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTecnicoSinGarantia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTecnicoSinGarantia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTecnicoSinGarantia.Location = new System.Drawing.Point(72, 16);
			this.cmbTecnicoSinGarantia.MaxDropDownItems = 30;
			this.cmbTecnicoSinGarantia.Name = "cmbTecnicoSinGarantia";
			this.cmbTecnicoSinGarantia.Size = new System.Drawing.Size(280, 21);
			this.cmbTecnicoSinGarantia.TabIndex = 240;
			this.cmbTecnicoSinGarantia.ValueMember = "idPersonal";
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Location = new System.Drawing.Point(16, 16);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(42, 17);
			this.label28.TabIndex = 239;
			this.label28.Text = "Técnico";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utpcGestionesTalleres
			// 
			this.utpcGestionesTalleres.Controls.Add(this.ultraDropDown1);
			this.utpcGestionesTalleres.Controls.Add(this.ultraGrid2);
			this.utpcGestionesTalleres.Location = new System.Drawing.Point(-10000, -10000);
			this.utpcGestionesTalleres.Name = "utpcGestionesTalleres";
			this.utpcGestionesTalleres.Size = new System.Drawing.Size(1222, 339);
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.ultraDataSource4;
			this.ultraDropDown1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 191;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 254;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraDropDown1.DisplayMember = "Descripcion";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(571, 153);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(256, 112);
			this.ultraDropDown1.TabIndex = 318;
			this.ultraDropDown1.ValueMember = "idGestionCliente";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "DTContactoCliente";
			this.ultraGrid2.DataSource = this.cdsDTContacto;
			appearance14.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance14;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "Band 0";
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 89;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 49;
			ultraGridColumn21.Header.Caption = "Fecha";
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn21.Width = 85;
			ultraGridColumn22.Header.Caption = "Gestión";
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 92;
			ultraGridColumn23.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.PromptChar = ' ';
			ultraGridColumn23.Width = 228;
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.PromptChar = ' ';
			ultraGridColumn24.Width = 70;
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 72;
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 99;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			ultraGridBand7.Header.Caption = "Gestión con cliente";
			ultraGridBand7.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 10F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(363, 9);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(496, 320);
			this.ultraGrid2.TabIndex = 317;
			// 
			// utpcGestionesClientes
			// 
			this.utpcGestionesClientes.Location = new System.Drawing.Point(-10000, -10000);
			this.utpcGestionesClientes.Name = "utpcGestionesClientes";
			this.utpcGestionesClientes.Size = new System.Drawing.Size(1222, 339);
			// 
			// btNotaCrédito
			// 
			this.btNotaCrédito.Enabled = false;
			this.btNotaCrédito.Location = new System.Drawing.Point(672, 704);
			this.btNotaCrédito.Name = "btNotaCrédito";
			this.btNotaCrédito.Size = new System.Drawing.Size(96, 25);
			this.btNotaCrédito.TabIndex = 314;
			this.btNotaCrédito.Text = "&Nota de Crédito";
			// 
			// btIngresoDT
			// 
			this.btIngresoDT.Enabled = false;
			this.btIngresoDT.Location = new System.Drawing.Point(576, 704);
			this.btIngresoDT.Name = "btIngresoDT";
			this.btIngresoDT.Size = new System.Drawing.Size(88, 25);
			this.btIngresoDT.TabIndex = 267;
			this.btIngresoDT.Text = "&Ingreso al DT";
			this.btIngresoDT.Click += new System.EventHandler(this.btIngresoDT_Click);
			// 
			// rdgTipoConsulta
			// 
			this.rdgTipoConsulta.Enabled = false;
			this.rdgTipoConsulta.Location = new System.Drawing.Point(136, 8);
			this.rdgTipoConsulta.Name = "rdgTipoConsulta";
			// 
			// rdgTipoConsulta.Properties
			// 
			this.rdgTipoConsulta.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipoConsulta.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipoConsulta.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Serial"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Factura"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Cédula")});
			this.rdgTipoConsulta.Size = new System.Drawing.Size(72, 56);
			this.rdgTipoConsulta.TabIndex = 214;
			this.rdgTipoConsulta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgTipoConsulta_KeyPress);
			this.rdgTipoConsulta.SelectedIndexChanged += new System.EventHandler(this.rdgTipoConsulta_SelectedIndexChanged);
			// 
			// btBuscar
			// 
			this.btBuscar.Enabled = false;
			this.btBuscar.Location = new System.Drawing.Point(568, 6);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 25);
			this.btBuscar.TabIndex = 213;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// txtBusca
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance21;
			this.txtBusca.Enabled = false;
			this.txtBusca.Location = new System.Drawing.Point(216, 7);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(344, 22);
			this.txtBusca.TabIndex = 212;
			this.txtBusca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusca_KeyPress);
			this.txtBusca.ValueChanged += new System.EventHandler(this.txtBusca_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 248);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(42, 16);
			this.lblBodega.TabIndex = 230;
			this.lblBodega.Text = "Técnico";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(776, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 229;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton7.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton7);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(840, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 228;
			this.dtFecha.Value = new System.DateTime(2014, 2, 25, 0, 0, 0, 0);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(888, 248);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(76, 16);
			this.label3.TabIndex = 234;
			this.label3.Text = "Observaciones";
			// 
			// txtObsIngreso
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObsIngreso.Appearance = appearance22;
			this.txtObsIngreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObsIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObsIngreso.Enabled = false;
			this.txtObsIngreso.Location = new System.Drawing.Point(976, 248);
			this.txtObsIngreso.Multiline = true;
			this.txtObsIngreso.Name = "txtObsIngreso";
			this.txtObsIngreso.Size = new System.Drawing.Size(240, 72);
			this.txtObsIngreso.TabIndex = 233;
			this.txtObsIngreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObsIngreso_KeyPress);
			this.txtObsIngreso.ValueChanged += new System.EventHandler(this.txtObsIngreso_ValueChanged);
			// 
			// txtIdIngreso
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance23.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdIngreso.Appearance = appearance23;
			this.txtIdIngreso.Enabled = false;
			this.txtIdIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdIngreso.FormatString = "";
			this.txtIdIngreso.Location = new System.Drawing.Point(1176, 704);
			this.txtIdIngreso.Name = "txtIdIngreso";
			this.txtIdIngreso.PromptChar = ' ';
			this.txtIdIngreso.Size = new System.Drawing.Size(40, 21);
			this.txtIdIngreso.TabIndex = 235;
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
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.DataSource = this.ultraDataSource1;
			this.cmbPersonal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 277;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(72, 248);
			this.cmbPersonal.MaxDropDownItems = 30;
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(296, 21);
			this.cmbPersonal.TabIndex = 238;
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPersonal_KeyPress);
			// 
			// lblNumero
			// 
			this.lblNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblNumero.Location = new System.Drawing.Point(1080, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(128, 22);
			this.lblNumero.TabIndex = 253;
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(384, 248);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(30, 16);
			this.label14.TabIndex = 261;
			this.label14.Text = "Daño";
			// 
			// txtDiagnostico
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiagnostico.Appearance = appearance24;
			this.txtDiagnostico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDiagnostico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiagnostico.Enabled = false;
			this.txtDiagnostico.Location = new System.Drawing.Point(432, 248);
			this.txtDiagnostico.Multiline = true;
			this.txtDiagnostico.Name = "txtDiagnostico";
			this.txtDiagnostico.Size = new System.Drawing.Size(448, 72);
			this.txtDiagnostico.TabIndex = 260;
			this.txtDiagnostico.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiagnostico_KeyPress);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.groupBox3);
			this.groupBox2.Enabled = false;
			this.groupBox2.Location = new System.Drawing.Point(0, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1224, 8);
			this.groupBox2.TabIndex = 262;
			this.groupBox2.TabStop = false;
			// 
			// groupBox3
			// 
			this.groupBox3.Enabled = false;
			this.groupBox3.Location = new System.Drawing.Point(0, 0);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(952, 8);
			this.groupBox3.TabIndex = 255;
			this.groupBox3.TabStop = false;
			// 
			// utcGarantia
			// 
			this.utcGarantia.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcGarantia.Appearance = appearance25;
			this.utcGarantia.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.utcGarantia.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcGarantia.Controls.Add(this.utpcConGarantia);
			this.utcGarantia.Controls.Add(this.utpcSinGarantia);
			this.utcGarantia.Controls.Add(this.utpcGestionesTalleres);
			this.utcGarantia.Controls.Add(this.utpcGestionesClientes);
			this.utcGarantia.Enabled = false;
			this.utcGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.utcGarantia.Location = new System.Drawing.Point(0, 336);
			this.utcGarantia.Name = "utcGarantia";
			this.utcGarantia.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcGarantia.Size = new System.Drawing.Size(1224, 360);
			this.utcGarantia.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcGarantia.TabIndex = 264;
			ultraTab1.TabPage = this.utpcConGarantia;
			ultraTab1.Text = "Gestión de Artículos con Garantía";
			ultraTab2.TabPage = this.utpcSinGarantia;
			ultraTab2.Text = "Gestión de Artículos sin Garantía";
			ultraTab3.TabPage = this.utpcGestionesTalleres;
			ultraTab3.Text = "Gestiones Taller";
			ultraTab4.TabPage = this.utpcGestionesClientes;
			ultraTab4.Text = "Gestiones Cliente";
			this.utcGarantia.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3,
																																													ultraTab4});
			this.utcGarantia.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1222, 339);
			// 
			// groupBox5
			// 
			this.groupBox5.Location = new System.Drawing.Point(0, 320);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1224, 8);
			this.groupBox5.TabIndex = 267;
			this.groupBox5.TabStop = false;
			// 
			// btNuevo
			// 
			this.btNuevo.Location = new System.Drawing.Point(8, 704);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(72, 25);
			this.btNuevo.TabIndex = 0;
			this.btNuevo.Text = "&Nuevo";
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.Enabled = false;
			this.btImprimir.Location = new System.Drawing.Point(248, 704);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(72, 25);
			this.btImprimir.TabIndex = 3;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Enabled = false;
			this.btCancelar.Location = new System.Drawing.Point(408, 704);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 25);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(488, 704);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 25);
			this.btSalir.TabIndex = 6;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.menuItem1,
																																								this.menuItem2});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Ingreso";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Egreso";
			// 
			// gbCliente
			// 
			this.gbCliente.Controls.Add(this.label39);
			this.gbCliente.Controls.Add(this.txtCelular);
			this.gbCliente.Controls.Add(this.txtidCliente);
			this.gbCliente.Controls.Add(this.label1);
			this.gbCliente.Controls.Add(this.label2);
			this.gbCliente.Controls.Add(this.label6);
			this.gbCliente.Controls.Add(this.label5);
			this.gbCliente.Controls.Add(this.txtTelefono);
			this.gbCliente.Controls.Add(this.txtDireccion);
			this.gbCliente.Controls.Add(this.txtCodigo);
			this.gbCliente.Controls.Add(this.txtNombre);
			this.gbCliente.Location = new System.Drawing.Point(8, 80);
			this.gbCliente.Name = "gbCliente";
			this.gbCliente.Size = new System.Drawing.Size(384, 160);
			this.gbCliente.TabIndex = 272;
			this.gbCliente.TabStop = false;
			this.gbCliente.Text = "Cliente";
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(200, 128);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(38, 16);
			this.label39.TabIndex = 275;
			this.label39.Text = "Celular";
			// 
			// txtCelular
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance26;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(256, 128);
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(112, 22);
			this.txtCelular.TabIndex = 274;
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			// 
			// txtidCliente
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance27.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidCliente.Appearance = appearance27;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(292, 32);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(80, 21);
			this.txtidCliente.TabIndex = 273;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(12, 128);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 272;
			this.label1.Text = "Teléfono";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(12, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 270;
			this.label2.Text = "Dirección";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(12, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(38, 16);
			this.label6.TabIndex = 268;
			this.label6.Text = "Código";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(12, 64);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(44, 16);
			this.label5.TabIndex = 266;
			this.label5.Text = "Nombre";
			// 
			// txtTelefono
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance28;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(76, 128);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 271;
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			// 
			// txtDireccion
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance29;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(76, 96);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(296, 22);
			this.txtDireccion.TabIndex = 269;
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			// 
			// txtCodigo
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance30;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(76, 32);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 267;
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			// 
			// txtNombre
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance31;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(76, 64);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 265;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// gbProducto
			// 
			this.gbProducto.Controls.Add(this.txtNumIngreso);
			this.gbProducto.Controls.Add(this.label38);
			this.gbProducto.Controls.Add(this.txtidDCFacVenta);
			this.gbProducto.Controls.Add(this.label30);
			this.gbProducto.Controls.Add(this.txtSerial);
			this.gbProducto.Controls.Add(this.txtIdSerial);
			this.gbProducto.Controls.Add(this.label13);
			this.gbProducto.Controls.Add(this.txtTiempo);
			this.gbProducto.Controls.Add(this.label9);
			this.gbProducto.Controls.Add(this.txtidArticulo);
			this.gbProducto.Controls.Add(this.label7);
			this.gbProducto.Controls.Add(this.label8);
			this.gbProducto.Controls.Add(this.txtCodArticulo);
			this.gbProducto.Controls.Add(this.txtArticulo);
			this.gbProducto.Controls.Add(this.cmbGarantia);
			this.gbProducto.Location = new System.Drawing.Point(400, 80);
			this.gbProducto.Name = "gbProducto";
			this.gbProducto.Size = new System.Drawing.Size(480, 160);
			this.gbProducto.TabIndex = 273;
			this.gbProducto.TabStop = false;
			this.gbProducto.Text = "Producto";
			// 
			// txtNumIngreso
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumIngreso.Appearance = appearance32;
			this.txtNumIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIngreso.Enabled = false;
			this.txtNumIngreso.FormatString = "";
			this.txtNumIngreso.Location = new System.Drawing.Point(400, 96);
			this.txtNumIngreso.Name = "txtNumIngreso";
			this.txtNumIngreso.PromptChar = ' ';
			this.txtNumIngreso.Size = new System.Drawing.Size(72, 22);
			this.txtNumIngreso.TabIndex = 289;
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(344, 99);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(47, 16);
			this.label38.TabIndex = 288;
			this.label38.Text = "Ingresos";
			// 
			// txtidDCFacVenta
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance33.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidDCFacVenta.Appearance = appearance33;
			this.txtidDCFacVenta.Enabled = false;
			this.txtidDCFacVenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidDCFacVenta.FormatString = "";
			this.txtidDCFacVenta.Location = new System.Drawing.Point(184, 128);
			this.txtidDCFacVenta.Name = "txtidDCFacVenta";
			this.txtidDCFacVenta.PromptChar = ' ';
			this.txtidDCFacVenta.Size = new System.Drawing.Size(48, 21);
			this.txtidDCFacVenta.TabIndex = 287;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(16, 99);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(32, 16);
			this.label30.TabIndex = 278;
			this.label30.Text = "Serial";
			// 
			// txtSerial
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance34;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Location = new System.Drawing.Point(64, 96);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(272, 22);
			this.txtSerial.TabIndex = 277;
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			this.txtSerial.ValueChanged += new System.EventHandler(this.txtSerial_ValueChanged);
			// 
			// txtIdSerial
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance35.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdSerial.Appearance = appearance35;
			this.txtIdSerial.Enabled = false;
			this.txtIdSerial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdSerial.FormatString = "";
			this.txtIdSerial.Location = new System.Drawing.Point(420, 32);
			this.txtIdSerial.Name = "txtIdSerial";
			this.txtIdSerial.PromptChar = ' ';
			this.txtIdSerial.Size = new System.Drawing.Size(48, 21);
			this.txtIdSerial.TabIndex = 276;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(240, 131);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(146, 16);
			this.label13.TabIndex = 275;
			this.label13.Text = "Tiempo restante de garantía";
			// 
			// txtTiempo
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance36;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Enabled = false;
			this.txtTiempo.FormatString = "";
			this.txtTiempo.Location = new System.Drawing.Point(400, 128);
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(72, 22);
			this.txtTiempo.TabIndex = 274;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.label9.Location = new System.Drawing.Point(16, 128);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(77, 23);
			this.label9.TabIndex = 272;
			this.label9.Text = "Garantía";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidArticulo
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance37.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidArticulo.Appearance = appearance37;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidArticulo.FormatString = "";
			this.txtidArticulo.Location = new System.Drawing.Point(364, 32);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(48, 21);
			this.txtidArticulo.TabIndex = 271;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(38, 16);
			this.label7.TabIndex = 270;
			this.label7.Text = "Código";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(16, 64);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(42, 16);
			this.label8.TabIndex = 268;
			this.label8.Text = "Artículo";
			// 
			// txtCodArticulo
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodArticulo.Appearance = appearance38;
			this.txtCodArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodArticulo.Enabled = false;
			this.txtCodArticulo.Location = new System.Drawing.Point(64, 29);
			this.txtCodArticulo.Name = "txtCodArticulo";
			this.txtCodArticulo.Size = new System.Drawing.Size(280, 22);
			this.txtCodArticulo.TabIndex = 269;
			this.txtCodArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodArticulo_KeyPress);
			// 
			// txtArticulo
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance39;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(64, 61);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(408, 22);
			this.txtArticulo.TabIndex = 267;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
			// 
			// cmbGarantia
			// 
			this.cmbGarantia.Enabled = false;
			this.cmbGarantia.Location = new System.Drawing.Point(104, 126);
			this.cmbGarantia.Name = "cmbGarantia";
			// 
			// cmbGarantia.Properties
			// 
			this.cmbGarantia.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.cmbGarantia.Properties.Appearance.Options.UseFont = true;
			this.cmbGarantia.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbGarantia.Properties.Items.AddRange(new object[] {
																																"SÍ",
																																"NO"});
			this.cmbGarantia.Size = new System.Drawing.Size(64, 26);
			this.cmbGarantia.TabIndex = 286;
			// 
			// gbProveedor
			// 
			this.gbProveedor.Controls.Add(this.txtGrupo);
			this.gbProveedor.Controls.Add(this.label37);
			this.gbProveedor.Controls.Add(this.label12);
			this.gbProveedor.Controls.Add(this.cmbFechaFactura);
			this.gbProveedor.Controls.Add(this.txtProveedor);
			this.gbProveedor.Controls.Add(this.txtNumFactura);
			this.gbProveedor.Controls.Add(this.label10);
			this.gbProveedor.Controls.Add(this.label11);
			this.gbProveedor.Location = new System.Drawing.Point(888, 80);
			this.gbProveedor.Name = "gbProveedor";
			this.gbProveedor.Size = new System.Drawing.Size(328, 160);
			this.gbProveedor.TabIndex = 274;
			this.gbProveedor.TabStop = false;
			this.gbProveedor.Text = "Proveedor";
			// 
			// txtGrupo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGrupo.Appearance = appearance40;
			this.txtGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupo.Enabled = false;
			this.txtGrupo.Location = new System.Drawing.Point(72, 128);
			this.txtGrupo.Name = "txtGrupo";
			this.txtGrupo.Size = new System.Drawing.Size(248, 22);
			this.txtGrupo.TabIndex = 259;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(8, 128);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(35, 16);
			this.label37.TabIndex = 258;
			this.label37.Text = "Grupo";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(8, 96);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(90, 16);
			this.label12.TabIndex = 257;
			this.label12.Text = "Fecha de Factura";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFechaFactura
			// 
			this.cmbFechaFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton8.Caption = "Today";
			this.cmbFechaFactura.DateButtons.Add(dateButton8);
			this.cmbFechaFactura.Enabled = false;
			this.cmbFechaFactura.Format = "dd/MM/yyyy";
			this.cmbFechaFactura.Location = new System.Drawing.Point(120, 96);
			this.cmbFechaFactura.Name = "cmbFechaFactura";
			this.cmbFechaFactura.NonAutoSizeHeight = 23;
			this.cmbFechaFactura.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaFactura.SpinButtonsVisible = true;
			this.cmbFechaFactura.TabIndex = 256;
			this.cmbFechaFactura.Value = ((object)(resources.GetObject("cmbFechaFactura.Value")));
			// 
			// txtProveedor
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor.Appearance = appearance41;
			this.txtProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor.Enabled = false;
			this.txtProveedor.Location = new System.Drawing.Point(72, 32);
			this.txtProveedor.Name = "txtProveedor";
			this.txtProveedor.Size = new System.Drawing.Size(248, 22);
			this.txtProveedor.TabIndex = 255;
			// 
			// txtNumFactura
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumFactura.Appearance = appearance42;
			this.txtNumFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumFactura.Enabled = false;
			this.txtNumFactura.Location = new System.Drawing.Point(72, 64);
			this.txtNumFactura.Name = "txtNumFactura";
			this.txtNumFactura.Size = new System.Drawing.Size(128, 22);
			this.txtNumFactura.TabIndex = 254;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 32);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(55, 16);
			this.label10.TabIndex = 253;
			this.label10.Text = "Proveedor";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(8, 64);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(56, 16);
			this.label11.TabIndex = 252;
			this.label11.Text = "N. Factura";
			// 
			// btGuardar
			// 
			this.btGuardar.Enabled = false;
			this.btGuardar.Location = new System.Drawing.Point(328, 704);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(72, 25);
			this.btGuardar.TabIndex = 4;
			this.btGuardar.Text = "&Grabar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
			// 
			// btEditar
			// 
			this.btEditar.Enabled = false;
			this.btEditar.Location = new System.Drawing.Point(168, 704);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(72, 25);
			this.btEditar.TabIndex = 2;
			this.btEditar.Text = "&Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btConsultar
			// 
			this.btConsultar.Location = new System.Drawing.Point(88, 704);
			this.btConsultar.Name = "btConsultar";
			this.btConsultar.Size = new System.Drawing.Size(72, 25);
			this.btConsultar.TabIndex = 1;
			this.btConsultar.Text = "&Consultar";
			this.btConsultar.Click += new System.EventHandler(this.btConsultar_Click);
			// 
			// lblNumeroEgreso
			// 
			this.lblNumeroEgreso.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblNumeroEgreso.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblNumeroEgreso.Location = new System.Drawing.Point(1080, 32);
			this.lblNumeroEgreso.Name = "lblNumeroEgreso";
			this.lblNumeroEgreso.Size = new System.Drawing.Size(128, 22);
			this.lblNumeroEgreso.TabIndex = 281;
			this.lblNumeroEgreso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.Location = new System.Drawing.Point(1008, 8);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(57, 16);
			this.label32.TabIndex = 282;
			this.label32.Text = "N. Ingreso";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.Location = new System.Drawing.Point(1008, 32);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(53, 16);
			this.label33.TabIndex = 283;
			this.label33.Text = "N. Egreso";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 16F);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(216, 32);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 29);
			this.lblEstado.TabIndex = 284;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// mnuTipoInforme
			// 
			this.mnuTipoInforme.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																									 this.mnuTipoInformeIngreso,
																																									 this.mnuTipoInformeEgreso});
			// 
			// mnuTipoInformeIngreso
			// 
			this.mnuTipoInformeIngreso.Index = 0;
			this.mnuTipoInformeIngreso.Text = "Ingreso";
			// 
			// mnuTipoInformeEgreso
			// 
			this.mnuTipoInformeEgreso.Index = 1;
			this.mnuTipoInformeEgreso.Text = "Egreso";
			// 
			// rdgOrigen
			// 
			this.rdgOrigen.EditValue = -1;
			this.rdgOrigen.Enabled = false;
			this.rdgOrigen.Location = new System.Drawing.Point(8, 8);
			this.rdgOrigen.Name = "rdgOrigen";
			// 
			// rdgOrigen.Properties
			// 
			this.rdgOrigen.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgOrigen.Properties.Appearance.Options.UseBackColor = true;
			this.rdgOrigen.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Point Technology"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Externo")});
			this.rdgOrigen.Size = new System.Drawing.Size(120, 40);
			this.rdgOrigen.TabIndex = 285;
			this.rdgOrigen.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdgOrigen_KeyPress);
			this.rdgOrigen.SelectedIndexChanged += new System.EventHandler(this.rdgOrigen_SelectedIndexChanged);
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(776, 32);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(41, 16);
			this.label36.TabIndex = 288;
			this.label36.Text = "Bodega";
			this.label36.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 180;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(840, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 289;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(8, 272);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(56, 16);
			this.label34.TabIndex = 291;
			this.label34.Text = "Accesorios";
			// 
			// txtAccesorios
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAccesorios.Appearance = appearance43;
			this.txtAccesorios.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAccesorios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAccesorios.Enabled = false;
			this.txtAccesorios.Location = new System.Drawing.Point(72, 272);
			this.txtAccesorios.Multiline = true;
			this.txtAccesorios.Name = "txtAccesorios";
			this.txtAccesorios.Size = new System.Drawing.Size(296, 48);
			this.txtAccesorios.TabIndex = 290;
			this.txtAccesorios.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAccesorios_KeyPress);
			// 
			// txtBodegaPredef
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance44.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtBodegaPredef.Appearance = appearance44;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBodegaPredef.FormatString = "";
			this.txtBodegaPredef.Location = new System.Drawing.Point(1128, 704);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(40, 21);
			this.txtBodegaPredef.TabIndex = 315;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 16F);
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(488, 32);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 29);
			this.lblUsuario.TabIndex = 316;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmSolDptoTecnico
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1224, 734);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.label34);
			this.Controls.Add(this.txtAccesorios);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.rdgOrigen);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.lblNumeroEgreso);
			this.Controls.Add(this.btConsultar);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.gbProveedor);
			this.Controls.Add(this.gbProducto);
			this.Controls.Add(this.gbCliente);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.utcGarantia);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtDiagnostico);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.txtIdIngreso);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtObsIngreso);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.rdgTipoConsulta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtBusca);
			this.Controls.Add(this.btIngresoDT);
			this.Controls.Add(this.btNotaCrédito);
			this.KeyPreview = true;
			this.Name = "frmSolDptoTecnico";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingreso Departamento Técnico";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSolDptoTecnico_KeyDown);
			this.Load += new System.EventHandler(this.frmSolDptoTecnico_Load);
			this.utpcConGarantia.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDTContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtGestionCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesIDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRetiro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbContactoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEnvio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroOrden)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTaller)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRevisado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados.Properties)).EndInit();
			this.utpcSinGarantia.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTecnicoSinGarantia)).EndInit();
			this.utpcGestionesTalleres.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoConsulta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObsIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiagnostico)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.utcGarantia)).EndInit();
			this.utcGarantia.ResumeLayout(false);
			this.gbCliente.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.gbProducto.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDCFacVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGarantia.Properties)).EndInit();
			this.gbProveedor.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOrigen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;
		public void HabilitaControles(Infragistics.Win.UltraWinTabControl.UltraTabControl Contenedor, bool Estado)
		{
			foreach (Control c in Contenedor.Controls) 
			{
				c.Enabled = Estado;
			}			
		}

		public void VaciaControles(GroupBox Contenedor)
		{
			foreach (Control c in Contenedor.Controls) 
			{
				if (c is Infragistics.Win.UltraWinEditors.UltraTextEditor) 
				{
					c.Text = "";
				}
			}
			this.txtIdIngreso.Value = 0;
			this.txtIdSerial.Value = 0;
			this.txtidCliente.Value = 0;
			this.txtidArticulo.Value = 0;
			this.txtNumIngreso.Value = 0;
			this.txtidDCFacVenta.Value = 0;
			this.txtTiempo.Value = 0;
		}

		private void frmSolDptoTecnico_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Today;

			this.cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where idDepartamento = 4 Order By Nombre Asc");
			this.cmbRevisado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where idDepartamento = 4 Order By Nombre Asc");
			this.cmbTaller.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idServicioTecnico, Nombre from ServicioTecnicoAut Order By Nombre Asc");
			this.cmbContacto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGestionCliente, Descripcion from DTGestionCliente Order By Descripcion Asc");
			this.cmbContactoCliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGestionCliente, Descripcion from DTGestionCliente Order By idGestionCliente Asc");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BodegaUsuario 51");

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void BuscaDatosSolicitud(string sSerial)
		{
			string sSQL = string.Format("Exec DTBuscarClientePorSerie '{0}'", sSerial);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidCliente.Value = dr.GetInt32(0);
				this.txtCodigo.Value = dr.GetString(1);
				this.txtNombre.Value = dr.GetString(2);
				this.txtDireccion.Value = dr.GetString(3);
				this.txtTelefono.Value = dr.GetString(4);
			}
			dr.Close();

			sSQL = string.Format("Exec DTBuscaArtPorSerial '{0}', {1}", sSerial, (int)this.txtidCliente.Value);
			SqlDataReader drArt = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drArt.Read();
			if (drArt.HasRows)
			{
				this.txtIdSerial.Value = drArt.GetInt32(0);
				this.txtidArticulo.Value = drArt.GetInt32(1);
				this.txtCodArticulo.Value = drArt.GetString(2);
				this.txtArticulo.Value = drArt.GetString(3);
				this.txtTiempo.Value = drArt.GetInt32(4);
				this.cmbGarantia.Text = drArt.GetString(5);
				this.txtidDCFacVenta.Value = drArt.GetInt32(6);
			}
			drArt.Close();

			sSQL = string.Format("Exec DTBuscaProvPorSerial '{0}'", sSerial);
			SqlDataReader drProv = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drProv.Read();
			if (drProv.HasRows)
			{
				this.txtProveedor.Value = drProv.GetString(0);
				this.txtNumFactura.Value = drProv.GetString(1);
				this.cmbFechaFactura.Value = drProv.GetDateTime(2);
				this.txtGrupo.Text = drProv.GetString(3);
			}
			drProv.Close();

			this.txtNumIngreso.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DTIngreso Where idSerial = {0}", (int)this.txtIdSerial.Value)) + 1;
		}

		private void HabilitaIngreso()
		{
			this.cmbGarantia.Enabled = true;
			this.cmbPersonal.Enabled = true;
			this.txtAccesorios.Enabled = true;
			this.txtDiagnostico.Enabled = true;
			this.txtObsIngreso.Enabled = true;
			this.dtEnvio.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			this.utcGarantia.Enabled = true;

			if (this.cmbGarantia.Text.ToString() == "NO")
			{
				this.utcGarantia.Tabs[0].Visible = false;
				this.utcGarantia.Tabs[1].Selected = true;
			}
			else
			{
				this.utcGarantia.Tabs[1].Visible = false;
				this.utcGarantia.Tabs[0].Selected = true;

				this.cmbTaller.Enabled = true;
				this.dtEnvio.Enabled = true;
			}

			this.btGuardar.Enabled = true; 
			this.cmbPersonal.Focus();
			if (bNuevo) this.lblEstado.Text = "PENDIENTE";
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.rdgTipoConsulta.EditValue == 0)
			{
				string sSQL = string.Format("Exec DTValidaBusquedaSolicitud '{0}'", this.txtBusca.Text.ToString());
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

				if (sMensaje.Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBusca.Focus();
					return;
				}

				BuscaDatosSolicitud(this.txtBusca.Text.ToString());
				this.txtSerial.Text = this.txtBusca.Text.ToString();

				HabilitaIngreso();
			}
			else if ((int)this.rdgTipoConsulta.EditValue == 1)
			{
				string sSQLFac = string.Format("Exec DTValidaBusquedaFactura '{0}'", this.txtBusca.Text.ToString());
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLFac, true);
				
				if (sMensaje.Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtBusca.Focus();
					return;
				}

				using (frmBuscaFacturaDT miBusqueda = new frmBuscaFacturaDT(this.txtBusca.Text.ToString(), (int)this.rdgTipoConsulta.EditValue))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						BuscaDatosSolicitud(miBusqueda.sSerial);
						this.txtSerial.Text = miBusqueda.sSerial;

						HabilitaIngreso();
					}
				}
			}
			else if ((int)this.rdgTipoConsulta.EditValue == 2)
			{
				using (frmBuscaFacturaDT miBusqueda = new frmBuscaFacturaDT(this.txtBusca.Text.ToString(), (int)this.rdgTipoConsulta.EditValue))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						BuscaDatosSolicitud(miBusqueda.sSerial);
						this.txtSerial.Text = miBusqueda.sSerial;

						HabilitaIngreso();
					}
				}
			}
		}

		private void txtBusca_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBusca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void frmSolDptoTecnico_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtIngreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtInforme.Focus();
		}

		private void dtInforme_Leave(object sender, System.EventArgs e)
		{
			if (this.dtInforme.Value == System.DBNull.Value) return;
			this.dtRetiro.CalendarInfo.MinDate = (DateTime)this.dtInforme.Value;
		}

		private void dtIngreso_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtIngreso_Leave(object sender, System.EventArgs e)
		{
			if (this.dtIngreso.Value == System.DBNull.Value) return;
			this.dtInforme.CalendarInfo.MinDate = (DateTime)this.dtIngreso.Value;
			DateTime FIngreso = (DateTime)this.dtIngreso.Value;
			this.dtInforme.Value = FIngreso.AddDays(2);
			this.dtRetiro.CalendarInfo.MinDate = (DateTime)this.dtInforme.Value;
		}

		private void btNuevo_Click(object sender, System.EventArgs e)
		{
			bNuevo = true;
			bEdicion = true;

			this.btConsultar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btImprimir.Enabled = false;
			this.btCancelar.Enabled = true;
			this.btNuevo.Enabled = false;

			this.cmbGarantia.Text = "";

			this.dtFecha.Value = DateTime.Today;
			this.rdgOrigen.EditValue = 0;
			
			if ((int)this.txtBodegaPredef.Value == 0) 
			{
				this.cmbBodega.Value = 1;
				this.cmbBodega.Enabled = true;
			}
			else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

			this.rdgOrigen.Enabled = true;
		}

		private void cmbPersonal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtAccesorios.Focus();
		}

		private void txtDiagnostico_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObsIngreso.Focus();
		}

		private void txtObsIngreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.utcGarantia.Tabs[0].Visible) this.cmbTaller.Focus();
				else if (this.utcGarantia.Tabs[1].Visible) this.cmbTecnicoSinGarantia.Focus();
			}
		}

		private void txtObsIngreso_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTaller_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtEnvio.Focus();
		}

		private void dtEnvio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNroOrden.Focus();
		}

		private void dtEnvio_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtNroOrden_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtIngreso.Focus();
		}

		private void dtInforme_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtInforme.Focus();
		}

		private void dtRetiro_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtInforme.Focus();
		}

		private void dtEnvio_Leave(object sender, System.EventArgs e)
		{
			if (this.dtEnvio.Value == System.DBNull.Value) return;
			this.dtIngreso.CalendarInfo.MinDate = (DateTime)this.dtEnvio.Value;
		}

		private void dtInforme_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.HabilitaControles(this.utcGarantia, true);
			this.rdgOrigen.EditValue = -1;
			this.rdgTipoConsulta.EditValue = 0;
			this.txtBusca.Text = "";
			this.rdgTipoConsulta.Enabled = false;
			this.rdgOrigen.Enabled = false;
			this.txtBusca.Enabled = false;
			this.btBuscar.Enabled = false;
			this.utcGarantia.Enabled = false;

			VaciaControles(this.gbCliente);
			VaciaControles(this.gbProducto);
			VaciaControles(this.gbProveedor);
	
			
			this.cmbPersonal.Value = System.DBNull.Value;
			this.txtAccesorios.Text = "";
			this.txtDiagnostico.Text = "";
			this.txtObsIngreso.Text = "";
			this.cmbPersonal.Enabled = false;
			this.cmbGarantia.Enabled = false;
			this.txtDiagnostico.Enabled = false;
			this.txtObsIngreso.Enabled = false;
			this.cmbGarantia.Text = "";
			this.lblEstado.Text = "";

			this.cmbTaller.Value = System.DBNull.Value;
			this.dtEnvio.Value = System.DBNull.Value;
			this.txtNroOrden.Text = "";
			this.dtIngreso.Value = System.DBNull.Value;
			this.dtInforme.Value = System.DBNull.Value;
			this.txtInforme.Text = "";
			this.dtRetiro.Value = System.DBNull.Value;
			this.cmbRevisado.Value = System.DBNull.Value;
			this.txtObservacionesIDT.Text = "";

			this.cmbFechaFactura.Value = System.DBNull.Value;

			this.btNuevo.Enabled = true;
			this.btConsultar.Enabled = true;
			this.btEditar.Enabled = false;
			this.btImprimir.Enabled = false;
			this.btGuardar.Enabled = false;
			this.btCancelar.Enabled = false;
			
			this.cmbBodega.Enabled = false;

			bNuevo = false;
			bEdicion = false;
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int)this.rdgOrigen.EditValue == 0)
			{
				if ((int)this.txtIdSerial.Value == 0)
				{
					MessageBox.Show("Ingrese un artículo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtBusca.Focus();
					return;
				}
				if (this.cmbPersonal.Text.ToString().Length == 0)
				{
					MessageBox.Show("Seleccione un técnico.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.cmbPersonal.Focus();
					return;
				}
				if (this.txtDiagnostico.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el diágnostico del artículo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtDiagnostico.Focus();
					return;
				}
			}
			#endregion Validacion

			int iEstado = 0;
			DateTime dtEnvio =  DateTime.Today;
			DateTime dtIngreso = DateTime.Today;
			DateTime dtInforme = DateTime.Today;
			DateTime dtRetiro = DateTime.Today;

			bool bFEnvio = false;
			bool bFIngreso = false;
			bool bFInforme = false;
			bool bFRetiro = false;

			if (this.dtEnvio.Value != System.DBNull.Value)
			{
				dtEnvio = (DateTime)this.dtEnvio.Value;
				bFEnvio = true;
			}
			if (this.dtIngreso.Value != System.DBNull.Value) 
			{
				dtIngreso = (DateTime)this.dtIngreso.Value;
				bFIngreso = true;
			}
			if (this.dtInforme.Value != System.DBNull.Value)
			{
				dtInforme = (DateTime)this.dtInforme.Value;
				bFInforme = true;
			}
			if (this.dtRetiro.Value != System.DBNull.Value) 
			{
				dtRetiro = (DateTime)this.dtRetiro.Value;
				bFRetiro = true;
			}
			
			int idTaller = 0;
			int iRevisadoPor = 0;
			
			if (this.cmbTaller.Text.ToString().Length > 0) idTaller = (int)this.cmbTaller.Value;
			if (this.cmbRevisado.Text.ToString().Length > 0) iRevisadoPor = (int)this.cmbRevisado.Value;

			#region Graba Ingreso
			if (bNuevo) this.lblNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 50");
			string sSQL = string.Format("Exec DTGuardaIngreso {0}, {1}, {2}, '{3}', {4}, {5}, '{6}', {7}, {8}, {9}, '{10}', '{11}', {12}, {13}, '{14}', '{15}', '{16}', '{17}', '{18}', '{19}', {20}, '{21}', {22}, {23}, {24}, {25}, '{26}'", 
				(int)this.txtIdIngreso.Value, (int)this.rdgOrigen.EditValue, (int)this.cmbBodega.Value, this.lblNumero.Text.ToString(), (int)this.txtidCliente.Value, (int)this.txtIdSerial.Value, this.cmbGarantia.Text.ToString(), 
				(int)this.txtTiempo.Value, (int)this.txtNumIngreso.Value, (int)this.cmbPersonal.Value, this.txtDiagnostico.Text.ToString(), this.txtObsIngreso.Text.ToString(), iEstado, idTaller, dtEnvio.ToString("yyyyMMdd"), 
				dtIngreso.ToString("yyyyMMdd"), this.txtNroOrden.Text.ToString(), dtInforme.ToString("yyyyMMdd"), dtRetiro.ToString("yyyyMMdd"), this.txtInforme.Text.ToString(), iRevisadoPor, this.txtObservacionesIDT.Text.ToString(),
				bFEnvio, bFIngreso, bFInforme, bFRetiro, this.txtAccesorios.Text.ToString());

			this.txtIdIngreso.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);			
			
			#region Crea Ingreso y Egreso de Bodega
			string sObservaciones = "";
			DateTime dtFecha = DateTime.Today;

			if (bNuevo)
			{
				dtFecha = (DateTime)this.dtFecha.Value;				
				sObservaciones = "INGRESO PARA REVISIÓN EN DEPARTAMENTO TÉCNICO.";

				string sSQLIEBod = string.Format("DTCreaINEGBodega {0}, {1}, {2}, '{3}', {4}, '{5}', '{6}', {7}, {8}, '{9}', {10}", 
					(int)this.txtidDCFacVenta.Value, (int)this.txtIdIngreso.Value, 9, dtFecha.ToString("yyyyMMdd"), (int)this.txtidCliente.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), 
					(int)this.txtIdSerial.Value, (int)this.cmbBodega.Value, sObservaciones, (int)this.rdgOrigen.EditValue);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLIEBod, true);
			}

			if (this.cmbTaller.Text.ToString().Length > 0 && this.dtEnvio.Text.ToString().Length == 0 && this.cmbGarantia.Text.ToString() == "SÍ")
			{
				dtFecha = (DateTime)this.dtFecha.Value;				
				sObservaciones = "SALIDA DEL DT HACIA EL TALLER AUTORIZADO " + this.cmbTaller.Text.ToString();

				string sSQLIEBod = string.Format("DTCreaINEGBodega {0}, {1}, {2}, '{3}', {4}, '{5}', '{6}', {7}, {8}, '{9}', {10}", 
					(int)this.txtIdIngreso.Value, 8, dtFecha.ToString("yyyyMMdd"), (int)this.txtidCliente.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), 
					(int)this.txtIdSerial.Value, (int)this.cmbBodega.Value, sObservaciones, (int)this.rdgOrigen.EditValue);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLIEBod, true);
			}
			#endregion Crea Ingreso y Egreso de Bodega

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idIngreso"].Value != System.DBNull.Value) dr.Cells["idIngreso"].Value = (int)this.txtIdIngreso.Value;
			}

			this.cdsDTContacto.Update();
			#endregion Graba Ingreso 

			this.btConsultar.Enabled = false;
			this.btEditar.Enabled = true;
			this.btImprimir.Enabled = true;
			this.btGuardar.Enabled = false;
			this.btCancelar.Enabled = true;
			this.rdgOrigen.Enabled = false;
			this.rdgTipoConsulta.Enabled = false;
			this.txtBusca.Enabled = false;
			this.btBuscar.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbGarantia.Enabled = false;
			this.cmbPersonal.Enabled = false;
			this.txtAccesorios.Enabled = false;
			this.txtDiagnostico.Enabled = false;
			this.txtObsIngreso.Enabled = false;
			this.utcGarantia.Enabled = false;
			this.cmbBodega.Enabled = false;

			bNuevo = false;
			bEdicion = false;
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			this.cmbPersonal.Enabled = true;
			this.txtAccesorios.Enabled = true;
			this.txtDiagnostico.Enabled = true;
			this.txtObsIngreso.Enabled = true;
			this.utcGarantia.Enabled = true;

			this.btConsultar.Enabled = false;
      this.btNuevo.Enabled = false;
			this.btEditar.Enabled = false;
			this.btImprimir.Enabled = false;
			this.btGuardar.Enabled = true;
			this.cmbGarantia.Enabled = true;

			iBodega = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 51, 3", true);
			if (iBodega == 0) this.cmbBodega.Enabled = true;

			bNuevo = false;
			bEdicion = true;
		}

		private void rdgTipoConsulta_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtBusca.SelectAll();
			this.txtBusca.Focus();
		}

		private void cdsDTContacto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsDTContacto.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void txtSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void rdgOrigen_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if ((int)this.rdgOrigen.EditValue == 0)
			{
				this.rdgTipoConsulta.Enabled = true;
				this.rdgTipoConsulta.EditValue = 0;
				this.txtBusca.Enabled = true;
				this.btBuscar.Enabled = true;
				this.txtBusca.Focus();

				this.txtCodigo.Enabled = false;
				this.txtNombre.Enabled = false;
				this.txtDireccion.Enabled = false;
				this.txtTelefono.Enabled = false;
				this.txtCelular.Enabled = false;

				this.txtCodArticulo.Enabled = false;
				this.txtArticulo.Enabled = false;
				this.txtSerial.Enabled = false;
			}
			else if ((int)this.rdgOrigen.EditValue == 1)
			{
				this.rdgTipoConsulta.Enabled = false;
				this.txtBusca.Enabled = false;
				this.btBuscar.Enabled = false;
				this.txtCodigo.Enabled = true;
				this.txtNombre.Enabled = true;
				this.txtDireccion.Enabled = true;
				this.txtTelefono.Enabled = true;
				this.txtCelular.Enabled = true;

				this.txtCodArticulo.Enabled = true;
				this.txtArticulo.Enabled = true;
				this.txtSerial.Enabled = true;

				this.cmbGarantia.Text = "NO";
				this.cmbGarantia.Enabled = false;
				this.txtCodigo.Focus();
			}
		}

		private void rdgOrigen_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.rdgTipoConsulta.Focus();
		}

		private void rdgTipoConsulta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtBusca.Focus();
		}

		private void btIngresoDT_Click(object sender, System.EventArgs e)
		{
			if (this.txtNroOrden.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Número de orden del Taller.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtNroOrden.Focus();
				return;
			}

			if (this.dtIngreso.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione la fecha de ingreso al Taller Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtIngreso.Focus();
				return;
			}

			if (this.dtInforme.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione la fecha de informe del Taller Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtInforme.Focus();
				return;
			}

			if (this.txtInforme.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el informe del Taller Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtInforme.Focus();
				return;
			}

			if (this.dtRetiro.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione la Fecha de Retiro del Taller Autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtInforme.Focus();
				return;
			}

			string sObservaciones = "";
			DateTime dtFecha = DateTime.Today;		

			if (this.cmbTaller.Text.ToString().Length > 0 && this.dtEnvio.Text.ToString().Length == 0 && this.cmbGarantia.Text.ToString() == "SÍ")
			{
				dtFecha = (DateTime)this.dtFecha.Value;				
				sObservaciones = "INGRESO AL DT DESDE EL TALLER AUTORIZADO " + this.cmbTaller.Text.ToString();

				string sSQLIEBod = string.Format("DTCreaINEGBodega {0}, {1}, {2}, '{3}', {4}, '{5}', '{6}', {7}, {8}, '{9}', {10}", 
					(int)this.txtIdIngreso.Value, 9, dtFecha.ToString("yyyyMMdd"), (int)this.txtidCliente.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), 
					(int)this.txtIdSerial.Value, (int)this.cmbBodega.Value, sObservaciones, (int)this.rdgOrigen.EditValue);
			}
		}

		private void cmbTaller_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtInforme_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtRetiro.Focus();
		}

		private void btConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmDTConsultaIngreso miBusqueda = new frmDTConsultaIngreso())
			{ 
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					/* Select i.idIngreso, i.idTipo, i.Bodega, i.Numero, i.Fecha, i.idCliente, cl.Ruc, cl.Nombre, cl.Direccion, cl.Telefono, 
						a.idArticulo, a.Codigo, a.Articulo, i.idSerial, s.Serial, i.Garantia, i.TiempoGarantia, i.Estado, 
						i.idPersonal, i.Diagnostico, i.Observacion, i.idServicioTecnico, i.Envio, i.NumeroOrden, 
						i.Ingreso, i.Informe, i.InformeDaños, i.Retiro, i.idPersonal1, i.Observacion1 */
					this.txtIdIngreso.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idIngreso"].Value;
					this.rdgOrigen.EditValue = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idTipo"].Value;
					this.cmbBodega.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Bodega"].Value; 
					this.lblNumero.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Numero"].Value.ToString();
					this.dtFecha.Value = (DateTime)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Fecha"].Value;	 
					this.txtidCliente.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idCliente"].Value;
					this.txtCodigo.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Ruc"].Value.ToString();	 
					this.txtNombre.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Nombre"].Value.ToString();	 
					this.txtDireccion.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Direccion"].Value.ToString();
					this.txtTelefono.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Telefono"].Value.ToString();	 
					this.txtidArticulo.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idArticulo"].Value;
					this.txtCodArticulo.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Codigo"].Value.ToString();
					this.txtArticulo.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Articulo"].Value.ToString();
					this.txtIdSerial.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idSerial"].Value;
					this.txtSerial.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Serial"].Value.ToString();
					this.cmbGarantia.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Garantia"].Value.ToString();
					this.txtTiempo.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["TiempoGarantia"].Value;
					this.txtNumIngreso.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["NumIngreso"].Value;

					#region Proveedor
					string sSQL = string.Format("Exec DTBuscaProvPorSerial '{0}'", this.txtSerial.Text.ToString());
					SqlDataReader drProv = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					drProv.Read();
					if (drProv.HasRows)
					{
						this.txtProveedor.Value = drProv.GetString(0);
						this.txtNumFactura.Value = drProv.GetString(1);
						this.cmbFechaFactura.Value = drProv.GetDateTime(2);
						this.txtGrupo.Text = drProv.GetString(3);
					}
					drProv.Close();
					#endregion Proveedor

					/* AQUI VA EL ESTADO */

					this.cmbPersonal.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idPersonal"].Value;
					this.txtAccesorios.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Accesorios"].Value.ToString();
					this.txtDiagnostico.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Diagnostico"].Value.ToString();
					this.txtObsIngreso.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Observacion"].Value.ToString();
					#region Garantia
					this.cmbTaller.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idServicioTecnico"].Value;
					if (miBusqueda.uGridDTIngreso.ActiveRow.Cells["Envio"].Value != System.DBNull.Value) this.dtEnvio.Value = (DateTime)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Envio"].Value;	  
					this.txtNroOrden.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["NumeroOrden"].Value.ToString();
					if (miBusqueda.uGridDTIngreso.ActiveRow.Cells["Ingreso"].Value != System.DBNull.Value) this.dtIngreso.Value = (DateTime)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Ingreso"].Value;
					if (miBusqueda.uGridDTIngreso.ActiveRow.Cells["Informe"].Value != System.DBNull.Value) this.dtInforme.Value = (DateTime)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Informe"].Value;	  
					this.txtInforme.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["InformeDaños"].Value.ToString();
					if (miBusqueda.uGridDTIngreso.ActiveRow.Cells["Retiro"].Value != System.DBNull.Value) this.dtRetiro.Value = (DateTime)miBusqueda.uGridDTIngreso.ActiveRow.Cells["Retiro"].Value;	  
					this.cmbRevisado.Value = (int)miBusqueda.uGridDTIngreso.ActiveRow.Cells["idPersonal1"].Value;
					this.txtObservacionesIDT.Text = miBusqueda.uGridDTIngreso.ActiveRow.Cells["Observacion1"].Value.ToString();
					#endregion Garantia

					this.btImprimir.Enabled = true;
					this.btEditar.Enabled = true;
					this.btCancelar.Enabled = true;
				}
			}
		}

		private void txtAccesorios_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDiagnostico.Focus();
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDireccion.Focus();
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtTelefono.Focus();
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCelular.Focus();
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCodArticulo.Focus();
		}

		private void txtCodArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtArticulo.Focus();
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtSerial.Focus();
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtAccesorios.Focus();
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}

