using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeguimientoAnadirNuevoMantenimientoLocal.
	/// </summary>
	public class frmSeguimientoAnadirNuevoMantenimientoLocal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Button btnSalir;
		private System.ComponentModel.IContainer components;
		int idResponsable = 0;
		int idAsignado = 0;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsMenu;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAsignado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipo;
		private System.Windows.Forms.TabPage tabPage4;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSeguimiento;
		private System.Windows.Forms.Button btnNuevo;
		int idEmpresa = 0;
		int idRequerimiento = 0;
		string Numero = "";
		bool bVer = false;
		bool bNuevo = false;
		int iNuevo = 1;
		int iEditar = 2;
		int iEliminar = 3;
		int iVer = 4;
		int getidSeguimiento = 0;
		int idSeguimientoPrincipal = 0;
		private System.Windows.Forms.GroupBox groupBox1;
		string Title = "";
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPersonal;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDocumento;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotaInterna;
		private System.Windows.Forms.Label lblNotaInterna;
		bool bRequerimientosDocumento = false;
		bool AsignaTecnico = true;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoUrgencia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoComplejidad;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.TabPage tabPage5;
		private System.Windows.Forms.TabPage tabPage6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaActual;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		public System.Windows.Forms.Label lblPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdRequerimientoComplejidad;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdRequerimientoUrgencia;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignado;
		private System.Windows.Forms.Label lblAsignado;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntregaAdm;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntregaTec;
		private System.Windows.Forms.TabControl tabDatosPrincipales;
		private System.Windows.Forms.Label lblFechaActual;
		private System.Windows.Forms.Label lblFechaEntregaAdm;
		private System.Windows.Forms.Label lblFechaEntregaTec;
		private System.Windows.Forms.GroupBox groupBox2;
		int iTipo = 0;
		public int idRequerimientoRol = 0;
		public string sRequerimientoRol = "";

		public frmSeguimientoAnadirNuevoMantenimientoLocal(int IdRequerimiento, int IdResponsable, string numero, int ITipo, int GetidSeguimiento)
		{
			idRequerimiento = IdRequerimiento;
			idResponsable = IdResponsable;
			Numero = numero;
			getidSeguimiento = GetidSeguimiento;
			iTipo = ITipo;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idempresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimientoAnadirNuevoMantenimientoLocal));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_seg");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombrePadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreSubHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlCompartido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlCompartido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoComplejidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoUrgencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoComplejidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoUrgencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsPersonal = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.udsMenu = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSeguimiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.lblNotaInterna = new System.Windows.Forms.Label();
			this.txtNotaInterna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoComplejidad = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoUrgencia = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.tabDatosPrincipales = new System.Windows.Forms.TabControl();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.lblFechaActual = new System.Windows.Forms.Label();
			this.dtFechaActual = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.cmdRequerimientoComplejidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.cmdRequerimientoUrgencia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbAsignado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.lblFechaEntregaAdm = new System.Windows.Forms.Label();
			this.dtFechaEntregaAdm = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.tabPage6 = new System.Windows.Forms.TabPage();
			this.dtFechaEntregaTec = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblFechaEntregaTec = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoComplejidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoUrgencia)).BeginInit();
			this.tabDatosPrincipales.SuspendLayout();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			this.tabPage5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmdRequerimientoComplejidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdRequerimientoUrgencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaAdm)).BeginInit();
			this.tabPage6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaTec)).BeginInit();
			this.SuspendLayout();
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// udsAsignado
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsAsignado.Band.Columns.AddRange(new object[] {
																														ultraDataColumn3,
																														ultraDataColumn4});
			// 
			// udsEmpresa
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsEmpresa.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn5,
																													 ultraDataColumn6});
			// 
			// udsPersonal
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.udsPersonal.Band.Columns.AddRange(new object[] {
																														ultraDataColumn7,
																														ultraDataColumn8});
			this.udsPersonal.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource1_CellDataRequested);
			// 
			// udsResponsable
			// 
			ultraDataColumn9.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn9,
																															 ultraDataColumn10});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(760, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 805;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Guardar seguimiento");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// udsMenu
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn12.DefaultValue = false;
			this.udsMenu.Band.Columns.AddRange(new object[] {
																												ultraDataColumn11,
																												ultraDataColumn12,
																												ultraDataColumn13,
																												ultraDataColumn14,
																												ultraDataColumn15});
			// 
			// udsDocumento
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
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
																														 ultraDataColumn26});
			// 
			// tabPage4
			// 
			this.tabPage4.Location = new System.Drawing.Point(0, 0);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.TabIndex = 0;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(760, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
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
			// ultraGrid1
			// 
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.TabIndex = 0;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Location = new System.Drawing.Point(0, 0);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 0;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-696, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2328, 8);
			this.groupBox1.TabIndex = 1037;
			this.groupBox1.TabStop = false;
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Location = new System.Drawing.Point(16, 160);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(824, 272);
			this.tabControl1.TabIndex = 1042;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.lblNotaInterna);
			this.tabPage1.Controls.Add(this.txtNotaInterna);
			this.tabPage1.Controls.Add(this.label6);
			this.tabPage1.Controls.Add(this.txtNota);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(816, 246);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Notas";
			// 
			// lblNotaInterna
			// 
			this.lblNotaInterna.Location = new System.Drawing.Point(8, 136);
			this.lblNotaInterna.Name = "lblNotaInterna";
			this.lblNotaInterna.Size = new System.Drawing.Size(80, 23);
			this.lblNotaInterna.TabIndex = 806;
			this.lblNotaInterna.Text = "Nota interna:";
			this.lblNotaInterna.Click += new System.EventHandler(this.label5_Click);
			// 
			// txtNotaInterna
			// 
			this.txtNotaInterna.AcceptsReturn = true;
			this.txtNotaInterna.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotaInterna.Appearance = appearance1;
			this.txtNotaInterna.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotaInterna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotaInterna.Location = new System.Drawing.Point(96, 128);
			this.txtNotaInterna.MaxLength = 2000;
			this.txtNotaInterna.Multiline = true;
			this.txtNotaInterna.Name = "txtNotaInterna";
			this.txtNotaInterna.Size = new System.Drawing.Size(680, 104);
			this.txtNotaInterna.TabIndex = 807;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(8, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 23);
			this.label6.TabIndex = 804;
			this.label6.Text = "* Nota:";
			// 
			// txtNota
			// 
			this.txtNota.AcceptsReturn = true;
			this.txtNota.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance2;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Location = new System.Drawing.Point(96, 16);
			this.txtNota.MaxLength = 2000;
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(680, 96);
			this.txtNota.TabIndex = 805;
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.ugdDocumento);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(816, 246);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Documentos";
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance3;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 150;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Width = 181;
			ultraGridColumn3.Header.Caption = "Seguimiento";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn3.Width = 79;
			ultraGridColumn4.Header.Caption = "Dirección archivo";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 476;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 135;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 135;
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 137;
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 154;
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 116;
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 73;
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 84;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn12.CellAppearance = appearance4;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			ultraGridColumn12.CellButtonAppearance = appearance5;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 32;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			ultraGridColumn13.CellAppearance = appearance6;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			ultraGridColumn13.CellButtonAppearance = appearance7;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 27;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn14.CellAppearance = appearance8;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn14.CellButtonAppearance = appearance9;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 23;
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
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(16, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(760, 208);
			this.ugdDocumento.TabIndex = 812;
			this.ugdDocumento.AfterRowUpdate += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowUpdate);
			this.ugdDocumento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowInsert);
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton_1);
			this.ugdDocumento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdDocumento_InitializeLayout_1);
			// 
			// udsRequerimientoComplejidad
			// 
			ultraDataColumn29.DataType = typeof(System.UInt32);
			this.udsRequerimientoComplejidad.Band.Columns.AddRange(new object[] {
																																						ultraDataColumn29,
																																						ultraDataColumn30});
			this.udsRequerimientoComplejidad.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource2_CellDataRequested);
			// 
			// udsRequerimientoUrgencia
			// 
			ultraDataColumn31.DataType = typeof(System.UInt32);
			this.udsRequerimientoUrgencia.Band.Columns.AddRange(new object[] {
																																				 ultraDataColumn31,
																																				 ultraDataColumn32});
			// 
			// tabDatosPrincipales
			// 
			this.tabDatosPrincipales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabDatosPrincipales.Controls.Add(this.tabPage3);
			this.tabDatosPrincipales.Controls.Add(this.tabPage5);
			this.tabDatosPrincipales.Controls.Add(this.tabPage6);
			this.tabDatosPrincipales.Location = new System.Drawing.Point(16, 16);
			this.tabDatosPrincipales.Name = "tabDatosPrincipales";
			this.tabDatosPrincipales.SelectedIndex = 0;
			this.tabDatosPrincipales.Size = new System.Drawing.Size(728, 112);
			this.tabDatosPrincipales.TabIndex = 1044;
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.lblFechaActual);
			this.tabPage3.Controls.Add(this.dtFechaActual);
			this.tabPage3.Controls.Add(this.cmbEstado);
			this.tabPage3.Controls.Add(this.lblEstado);
			this.tabPage3.Controls.Add(this.label12);
			this.tabPage3.Controls.Add(this.txtNumero);
			this.tabPage3.Controls.Add(this.cmbPersonal);
			this.tabPage3.Controls.Add(this.lblPersonal);
			this.tabPage3.Controls.Add(this.cmbResponsable);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(720, 86);
			this.tabPage3.TabIndex = 0;
			this.tabPage3.Tag = "0";
			this.tabPage3.Text = "Datos principales";
			// 
			// lblFechaActual
			// 
			this.lblFechaActual.Location = new System.Drawing.Point(352, 48);
			this.lblFechaActual.Name = "lblFechaActual";
			this.lblFechaActual.Size = new System.Drawing.Size(96, 24);
			this.lblFechaActual.TabIndex = 1074;
			this.lblFechaActual.Text = "* Fecha actual:";
			// 
			// dtFechaActual
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaActual.Appearance = appearance14;
			this.dtFechaActual.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaActual.DateButtons.Add(dateButton1);
			this.dtFechaActual.Enabled = false;
			this.dtFechaActual.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaActual.Location = new System.Drawing.Point(448, 48);
			this.dtFechaActual.Name = "dtFechaActual";
			this.dtFechaActual.NonAutoSizeHeight = 23;
			this.dtFechaActual.Size = new System.Drawing.Size(240, 21);
			this.dtFechaActual.SpinButtonsVisible = true;
			this.dtFechaActual.TabIndex = 1073;
			this.dtFechaActual.Value = ((object)(resources.GetObject("dtFechaActual.Value")));
			// 
			// cmbEstado
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance15;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "Descripción";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 240;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(104, 48);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(240, 21);
			this.cmbEstado.TabIndex = 1072;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.TextChanged += new System.EventHandler(this.cmbEstado_TextChanged);
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstado_InitializeLayout);
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(8, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.TabIndex = 1071;
			this.lblEstado.Text = "* Estado:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(360, 16);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(48, 16);
			this.label12.TabIndex = 1070;
			this.label12.Text = "Número:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance16;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(448, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.ReadOnly = true;
			this.txtNumero.Size = new System.Drawing.Size(240, 21);
			this.txtNumero.TabIndex = 1069;
			// 
			// cmbPersonal
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPersonal.Appearance = appearance17;
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.udsPersonal;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(104, 16);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(240, 21);
			this.cmbPersonal.TabIndex = 1067;
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// lblPersonal
			// 
			this.lblPersonal.Location = new System.Drawing.Point(8, 16);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.TabIndex = 1066;
			this.lblPersonal.Text = "* Personal:";
			// 
			// cmbResponsable
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance18;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(696, 16);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(24, 21);
			this.cmbResponsable.TabIndex = 1068;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.Visible = false;
			// 
			// tabPage5
			// 
			this.tabPage5.Controls.Add(this.cmdRequerimientoComplejidad);
			this.tabPage5.Controls.Add(this.label4);
			this.tabPage5.Controls.Add(this.cmdRequerimientoUrgencia);
			this.tabPage5.Controls.Add(this.label2);
			this.tabPage5.Controls.Add(this.cmbAsignado);
			this.tabPage5.Controls.Add(this.lblAsignado);
			this.tabPage5.Controls.Add(this.lblFechaEntregaAdm);
			this.tabPage5.Controls.Add(this.dtFechaEntregaAdm);
			this.tabPage5.Location = new System.Drawing.Point(4, 22);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(720, 86);
			this.tabPage5.TabIndex = 1;
			this.tabPage5.Tag = "1";
			this.tabPage5.Text = "Asignación";
			// 
			// cmdRequerimientoComplejidad
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdRequerimientoComplejidad.Appearance = appearance19;
			this.cmdRequerimientoComplejidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmdRequerimientoComplejidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdRequerimientoComplejidad.DataSource = this.udsRequerimientoComplejidad;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 240;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmdRequerimientoComplejidad.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmdRequerimientoComplejidad.DisplayMember = "Descripcion";
			this.cmdRequerimientoComplejidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdRequerimientoComplejidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdRequerimientoComplejidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdRequerimientoComplejidad.Location = new System.Drawing.Point(448, 16);
			this.cmdRequerimientoComplejidad.Name = "cmdRequerimientoComplejidad";
			this.cmdRequerimientoComplejidad.Size = new System.Drawing.Size(240, 21);
			this.cmdRequerimientoComplejidad.TabIndex = 1063;
			this.cmdRequerimientoComplejidad.ValueMember = "idRequerimientoComplejidad";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(352, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 23);
			this.label4.TabIndex = 1061;
			this.label4.Text = "* Complejidad:";
			// 
			// cmdRequerimientoUrgencia
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdRequerimientoUrgencia.Appearance = appearance20;
			this.cmdRequerimientoUrgencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmdRequerimientoUrgencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdRequerimientoUrgencia.DataSource = this.udsRequerimientoUrgencia;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.Caption = "Descripción";
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 241;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmdRequerimientoUrgencia.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmdRequerimientoUrgencia.DisplayMember = "Descripcion";
			this.cmdRequerimientoUrgencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdRequerimientoUrgencia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdRequerimientoUrgencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdRequerimientoUrgencia.Location = new System.Drawing.Point(104, 16);
			this.cmdRequerimientoUrgencia.Name = "cmdRequerimientoUrgencia";
			this.cmdRequerimientoUrgencia.Size = new System.Drawing.Size(240, 21);
			this.cmdRequerimientoUrgencia.TabIndex = 1059;
			this.cmdRequerimientoUrgencia.ValueMember = "idRequerimientoUrgencia";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 16);
			this.label2.Name = "label2";
			this.label2.TabIndex = 1058;
			this.label2.Text = "* Urgencia:";
			// 
			// cmbAsignado
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAsignado.Appearance = appearance21;
			this.cmbAsignado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAsignado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignado.DataSource = this.udsAsignado;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 284;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbAsignado.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbAsignado.DisplayMember = "Nombre";
			this.cmbAsignado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAsignado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignado.Location = new System.Drawing.Point(448, 48);
			this.cmbAsignado.Name = "cmbAsignado";
			this.cmbAsignado.Size = new System.Drawing.Size(240, 21);
			this.cmbAsignado.TabIndex = 1057;
			this.cmbAsignado.ValueMember = "idPersonal";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(352, 48);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.TabIndex = 1056;
			this.lblAsignado.Text = "* Asignado:";
			// 
			// lblFechaEntregaAdm
			// 
			this.lblFechaEntregaAdm.Location = new System.Drawing.Point(8, 48);
			this.lblFechaEntregaAdm.Name = "lblFechaEntregaAdm";
			this.lblFechaEntregaAdm.Size = new System.Drawing.Size(96, 24);
			this.lblFechaEntregaAdm.TabIndex = 1062;
			this.lblFechaEntregaAdm.Text = "* Fecha entrega:";
			// 
			// dtFechaEntregaAdm
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntregaAdm.Appearance = appearance22;
			this.dtFechaEntregaAdm.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaEntregaAdm.DateButtons.Add(dateButton2);
			this.dtFechaEntregaAdm.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaEntregaAdm.Location = new System.Drawing.Point(104, 48);
			this.dtFechaEntregaAdm.Name = "dtFechaEntregaAdm";
			this.dtFechaEntregaAdm.NonAutoSizeHeight = 23;
			this.dtFechaEntregaAdm.Size = new System.Drawing.Size(240, 21);
			this.dtFechaEntregaAdm.SpinButtonsVisible = true;
			this.dtFechaEntregaAdm.TabIndex = 1060;
			this.dtFechaEntregaAdm.Value = ((object)(resources.GetObject("dtFechaEntregaAdm.Value")));
			// 
			// tabPage6
			// 
			this.tabPage6.Controls.Add(this.dtFechaEntregaTec);
			this.tabPage6.Controls.Add(this.lblFechaEntregaTec);
			this.tabPage6.Location = new System.Drawing.Point(4, 22);
			this.tabPage6.Name = "tabPage6";
			this.tabPage6.Size = new System.Drawing.Size(720, 86);
			this.tabPage6.TabIndex = 2;
			this.tabPage6.Tag = "2";
			this.tabPage6.Text = "Datos del técnico";
			// 
			// dtFechaEntregaTec
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntregaTec.Appearance = appearance23;
			this.dtFechaEntregaTec.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaEntregaTec.DateButtons.Add(dateButton3);
			this.dtFechaEntregaTec.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaEntregaTec.Location = new System.Drawing.Point(104, 16);
			this.dtFechaEntregaTec.Name = "dtFechaEntregaTec";
			this.dtFechaEntregaTec.NonAutoSizeHeight = 23;
			this.dtFechaEntregaTec.Size = new System.Drawing.Size(240, 21);
			this.dtFechaEntregaTec.SpinButtonsVisible = true;
			this.dtFechaEntregaTec.TabIndex = 1060;
			this.dtFechaEntregaTec.Value = ((object)(resources.GetObject("dtFechaEntregaTec.Value")));
			// 
			// lblFechaEntregaTec
			// 
			this.lblFechaEntregaTec.Location = new System.Drawing.Point(8, 16);
			this.lblFechaEntregaTec.Name = "lblFechaEntregaTec";
			this.lblFechaEntregaTec.Size = new System.Drawing.Size(112, 24);
			this.lblFechaEntregaTec.TabIndex = 1061;
			this.lblFechaEntregaTec.Text = "* Fecha entrega:";
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(8, 136);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(8, 8);
			this.groupBox2.TabIndex = 1045;
			this.groupBox2.TabStop = false;
			// 
			// frmSeguimientoAnadirNuevoMantenimientoLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(856, 446);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.tabDatosPrincipales);
			this.Name = "frmSeguimientoAnadirNuevoMantenimientoLocal";
			this.Load += new System.EventHandler(this.frmSeguimientoAnadirNuevoMantenimientoLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoComplejidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoUrgencia)).EndInit();
			this.tabDatosPrincipales.ResumeLayout(false);
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			this.tabPage5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmdRequerimientoComplejidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdRequerimientoUrgencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaAdm)).EndInit();
			this.tabPage6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaTec)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables

		bool bEdicion = false;
		bool bImpreso = false;

		private Acceso miAcceso;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;		
		#endregion Variables

		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";
		string FechaAsignacion;
		DateTime dtFecha = DateTime.Now;
		int idUsuarioInicio = 0;

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ControlCampos()
		{
			if(bNuevo)
			{
				if(idRequerimientoRol == frmRequerimiento.iRequerimientoRolTecnico)
				{
					this.cmbAsignado.Enabled = false;
					this.dtFechaEntregaAdm.Enabled = false;
					this.cmdRequerimientoUrgencia.Enabled = false;
					this.cmdRequerimientoComplejidad.Enabled = false;
					string dtFechaEntregaTec = "";
					try{dtFechaEntregaTec = Convert.ToDateTime(this.dtFechaEntregaTec.Value).ToString("yyyyMMdd HH:mm:ss");}
					catch(Exception ex){string sms = ex.Message;}
					if(dtFechaEntregaTec != "")
					{
						this.dtFechaEntregaTec.Enabled = false;
					}
				}
			}

			if(idRequerimientoRol == frmRequerimiento.iRequerimientoRolUsuario)
			{
				this.tabPage5.Parent = null;
				this.tabPage6.Parent = null;
			}

			if(bVer)
			{
				this.cmbAsignado.Enabled = false;
				this.dtFechaEntregaAdm.Enabled = false;
				this.cmdRequerimientoUrgencia.Enabled = false;
				this.cmdRequerimientoComplejidad.Enabled = false;
				this.dtFechaEntregaTec.Enabled = false;
			}

		}

		private void frmSeguimientoAnadirNuevoMantenimientoLocal_Load(object sender, System.EventArgs e)
		{
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();

			try
			{
				if(iTipo == iNuevo)
				{
					bNuevo = true;
					Title = "Nuevo seguimiento";
				}
				if(iTipo == iVer)
				{
					bVer = true;
					Title = "Ver seguimiento";
				}
				this.Text = Title;
				//Carga combo empresa
				SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'");
				SqlCommand oCmd = new SqlCommand();
				oCmd.Connection = oConexion;
				oCmd.CommandType = CommandType.Text;
				//Consulta el id de la tabla inicio, por el nombre de usuario que se encuentra ingresado
				SqlCommand oCmdEmp = new SqlCommand(string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario), oConexion);
				oConexion.Open();
				idUsuarioInicio = (int)oCmdEmp.ExecuteScalar();
				oCmd.CommandText = string.Format("Select e.idEmpresa, e.Empresa From Empresas e Inner Join EmpresasUsuarios eu On eu.idEmpresa = e.idEmpresa Where eu.idUsuario = {0} Order By e.Empresa", idUsuarioInicio);
				SqlDataAdapter oDA = new SqlDataAdapter(oCmd.CommandText, oConexion);
				DataSet oDS = new DataSet();
				oDA.Fill(oDS);
				if(bNuevo)
				{
					int getidPersonal = 0;
					getidPersonal = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("Exec RequerimientoResponsableConsultar {0}", idUsuarioInicio),true);
					//Llenar el combo del personal, es el id de usuario logeado
					this.cmbPersonal.Value = getidPersonal;
				}
				//CargarResponsable
				this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
				this.cmbResponsable.Value = idResponsable;
				//Cargar la persona que hace el seguimiento
				this.cmbPersonal.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
				DateTime GetFechaAsignacion =  Funcion.dtEjecucion(frmRequerimiento.sconexionPoint, string.Format("SELECT FechaAsignacion FROM Requerimiento WHERE idRequerimiento = {0}",idRequerimiento));
				FechaAsignacion = GetFechaAsignacion.ToString("yyyyMMdd");
				//Cargar combos asignados
				this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}, {3}", idResponsable, idAsignado, AsignaTecnico, frmRequerimiento.iRequerimientoControlMantenimientoLocales));
				ControlNuevo();
				ControlVer();
				//Controles de permisos
				RequerimientoRolPermisoCon();
				//Requerimiento urgencia
				cmdRequerimientoUrgencia.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, "Exec RequerimientoUrgenciaCon");
				//Requerimiento complejidad
				cmdRequerimientoComplejidad.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, "Exec RequerimientoComplejidadCon");
				//Número
				string squery = String.Format("Exec RequerimientoModel '{0}', {1}", "Numero", idRequerimiento);
				this.txtNumero.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				//Nombre del label personal
				lblPersonalDesignacion();
				//Fecha actual, para comparativa
				//Cargar la fecha de ingreso
				this.dtFechaActual.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				//this.dtFechaTentativa.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				DateTime dtFechaActual = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
				this.dtFechaActual.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				//this.dtFechaTentativa.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				this.dtFechaActual.CalendarInfo.MaxDate = dtFechaActual;
				//this.dtFechaTentativa.CalendarInfo.MaxDate = dtFecha;
				this.dtFechaActual.Value = "01/" + dtFechaActual.Month.ToString() + "/" + dtFechaActual.Year.ToString();
				this.dtFechaActual.Value = dtFechaActual;	
				CargarDataSeguimientoNuevo();
				//Controls por roles
				ControlCampos();
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}

		}

		private void CargarDataSeguimientoNuevo()
		{
			int idSeguimientoCargaData = 0;
			string squery = "";
				if(bNuevo)
				{
					//Carga fecha entrega tecnico
					//Sacar el idseguimiento principal del requerimiento.
					squery = String.Format("Select idSeguimiento From Seguimiento Where idRequerimiento = {0} And Principal = 1", idRequerimiento);
					idSeguimientoCargaData = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery, false);
				}
				else
				{
					idSeguimientoCargaData = getidSeguimiento;
				}

				DateTime fechaDefault = new DateTime();

				try
				{
					squery = String.Format("EXEC SeguimientoModel'{0}' ,{1}","FechaEntregaTec", idSeguimientoCargaData);
					DateTime getdtFechaEntregaTec = Funcion.dtEjecucion(frmRequerimiento.sconexionPoint,squery, true);
					if(fechaDefault != getdtFechaEntregaTec)
					{
							this.dtFechaEntregaTec.Value = getdtFechaEntregaTec;
					}

				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}

				try
				{
					squery = String.Format("EXEC SeguimientoModel'{0}' ,{1}","FechaEntregaAdm", idSeguimientoCargaData);
					DateTime getdtFechaEntregaAdm = Funcion.dtEjecucion(frmRequerimiento.sconexionPoint,squery, true);
					if(fechaDefault != getdtFechaEntregaAdm)
					{
						this.dtFechaEntregaAdm.Value = getdtFechaEntregaAdm;
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}

				try
				{
					squery = String.Format("EXEC SeguimientoModel'{0}' ,{1}","idRequerimientoUrgencia", idSeguimientoCargaData);
					int getcmdRequerimientoUrgencia = Funcion.iEjecucion(frmRequerimiento.sconexionPoint,squery, false);
					if(getcmdRequerimientoUrgencia > 0)
					{
						this.cmdRequerimientoUrgencia.Value = getcmdRequerimientoUrgencia;
					}
					
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}

				try
				{
					squery = String.Format("EXEC SeguimientoModel'{0}' ,{1}","idRequerimientoComplejidad", idSeguimientoCargaData);
					int getcmdRequerimientoComplejidad = Funcion.iEjecucion(frmRequerimiento.sconexionPoint,squery, false);
					if(getcmdRequerimientoComplejidad > 0)
					{
						this.cmdRequerimientoComplejidad.Value  = getcmdRequerimientoComplejidad;
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}
		}
		private void lblPersonalDesignacion()
		{
			if(bNuevo)
			{
				//Ubicar si es administrador, técnico o usuario.
				this.lblPersonal.Text = sRequerimientoRol + ": ";
			}
			else
			{
				//Cual es el rol del usuario a ver
				int idRequerimientoRol = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}", this.cmbPersonal.Value), false);
				//Label personal
				string squery = String.Format("Exec RequerimientoRolModel '{0}', {1}", "Subtitulo", idRequerimientoRol);
				string sRequerimientoRol = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				this.lblPersonal.Text = sRequerimientoRol + ": ";
			}
		
		}

		private void RequerimientoRolPermisoCon()
		{
			try
			{
				//Permiso Asignación
				bool bAsignacion = Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0},'{1}', {2}", 
					(int)this.cmbPersonal.Value,
					frmRequerimiento.iAsignar,
					frmRequerimiento.iRequerimientoControlMantenimientoLocales
					), false); 
				if(bAsignacion == false)
				{
					this.cmbAsignado.Visible = false;
					this.lblAsignado.Visible = false;
				}
				//Permiso Valores técnicos
				bool bvalorestecnicos =
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0},'{1}', {2}", 
					(int)this.cmbPersonal.Value,
					frmRequerimiento.iValoresTecnicos,
					frmRequerimiento.iRequerimientoControlMantenimientoLocales
					), false); 

				if(bvalorestecnicos == false)
				{
					this.lblNotaInterna.Visible = false;
					this.txtNotaInterna.Visible = false;
				}


			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				return;
			}
			
		}


		private void ControlVer()
		{
			if(bVer)
			{
				this.btnNuevo.Enabled = false;
				this.cmbResponsable.Enabled = false;
				this.cmbPersonal.Enabled = false;
				this.txtNota.ReadOnly = true;
				this.btnGuardar.Enabled = false;
				this.cmbEstado.Enabled = false;
				this.txtNota.Value = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("Select Notas From Seguimiento where idSeguimiento = {0}", getidSeguimiento));
				//Llenar el combo del personal, es el id de usuario logeado
				this.cmbPersonal.Value = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoSeguimientoPersonalCon {0}", getidSeguimiento),false);
				//Cargar EstadoRequerimiento
				this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoEstado"));
				//Cargar el estado actual del requerimiento
				int Estado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint,String.Format("Select idEstado From Seguimiento where idSeguimiento = {0}", getidSeguimiento),false);
				this.cmbEstado.Value = Estado; 
				//Consultar documento
				RequerimientoDocumentoIdSeguimiento();
				//Cargar id Asignado
				this.cmbAsignado.Value = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("Select idAsignado From Seguimiento where idSeguimiento = {0}", getidSeguimiento),false);
				//Cargar Nota interna
				this.txtNotaInterna.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("Select NotaInterna From Seguimiento where idSeguimiento = {0}", getidSeguimiento));
				this.cmbAsignado.Enabled = false;
				this.txtNotaInterna.ReadOnly = true;
			}
		}

		private void ControlNuevo()
		{
			if(bNuevo)
			{
				//Cargar EstadoRequerimiento
				this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoRolEstadoCon {0}, {1}", (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales));
				//Cargar el estado actual del requerimiento
				int Estado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint,String.Format("SELECT idEstado FROM Requerimiento where idRequerimiento = {0}", idRequerimiento),false);
				this.cmbEstado.Value = Estado; 
				bool Activo = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoRolEstadoActivo {0}, {1}, {2}", (int)this.cmbPersonal.Value, Estado, frmRequerimiento.iRequerimientoControlMantenimientoLocales),false);
				this.cmbEstado.Enabled = Activo;
				int idPersonal = Funcion.iEjecucion(frmRequerimiento.sconexionPoint,String.Format("SELECT idPersonal FROM Seguimiento where idRequerimiento = {0} AND Principal = 1", idRequerimiento),false);
				this.cmbAsignado.Value = idPersonal; 
				//Consultar documento
				RequerimientoDocumentoCon(1);
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		public bool ValidarAdmnistrador()
		{
			if(idRequerimientoRol == frmRequerimiento.iRequerimientoRolAdministrador)
			{
				//Validar si existe una asignación y es administrador, está debe ubicar, el asignado, urgencias, complegidad, fecha de entrega
				if((int)this.cmbEstado.Value == frmRequerimiento.iEstadoRequerimientoAsignado)
				{
					if (!Validacion.vbComboVacio(this.cmbAsignado, "Seleccione la persona asignada. ")) 
					{
						this.tabDatosPrincipales.SelectedIndex = 1;
						return false;
					}
					if (!Validacion.vbComboVacio(this.cmdRequerimientoUrgencia, "Seleccione la urgencia. ")) 
					{
						this.tabDatosPrincipales.SelectedIndex = 1;
						return false;
					}
					if (!Validacion.vbComboVacio(this.cmdRequerimientoComplejidad, "Seleccione la complejidad. ")) 
					{
						this.tabDatosPrincipales.SelectedIndex = 1;
						return false;
					}
					if (!Validacion.vbValidaFechasDesdeHasta(this.dtFechaActual, this.dtFechaEntregaAdm, "fecha actual", "fecha de entrega")) 
					{
						this.tabDatosPrincipales.SelectedIndex = 1;
						return false;
					}
				}
			}
			return true;
		}

		public bool ValidarUsuario()
		{
			//Validar si existe una asignación y es administrador, está debe ubicar, el asignado, urgencias, complegidad, fecha de entrega
			if(idRequerimientoRol == frmRequerimiento.iRequerimientoRolUsuario)
			{
				
			}
			return true;
		}

		public bool ValidarTecnico()
		{
			//Validar si existe una asignación y es administrador, está debe ubicar, el asignado, urgencias, complegidad, fecha de entrega
			if(idRequerimientoRol == frmRequerimiento.iRequerimientoRolTecnico)
			{
				//Se valida la fecha cuando el estado seleccionado este en proceso
				if((int)this.cmbEstado.Value == frmRequerimiento.iEstadoRequerimientoEnProceso)
				{
					if (!Validacion.vbValidaFechasDesdeHasta(this.dtFechaActual, this.dtFechaEntregaTec, "fecha actual", "fecha de entrega")) 
					{
						this.tabDatosPrincipales.SelectedIndex = 2;
						return false;
					}
				}
			}
			return true;
		}


		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region valida Campos Generales 
			//if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione el responsable ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) 
			{
				this.tabDatosPrincipales.SelectedIndex = 0;
				return;
			}
			if (!Validacion.vbTexto(this.txtNota, 20, 2000, "Ingrese una nota")) return;
			if(!ValidarAdmnistrador()) return;
			if(!ValidarTecnico()) return;
			if(!ValidarUsuario()) return;
			#endregion valida Campos Generales 

			#region Valida documentos
			if(bRequerimientosDocumento)
			{
				// Validación si se ha ingresado al menos un documentos eso solo sería de hacer un select en la tabla documentosrequerimiento con el estado 1
				int ugdDocumentotamano = this.ugdDocumento.Rows.Count;
				if (ugdDocumentotamano == 0)
				{
					MessageBox.Show(string.Format("Proceso Finalizado, se debe cargar mínimo un documento"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (ugdDocumentotamano > 3)
				{
					MessageBox.Show(string.Format("Proceso Finalizado, se debe cargar máximo 3 documentos"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida documentos 

			Resultado resultado = new Resultado();

			#region Validar politicas
			resultado = frmRequerimiento.RequerimientoPoliticaVal(1,(int)this.cmbEstado.Value, idRequerimiento, (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
			if(resultado.bResultado == false)
			{
				//string spolitica1 = frmRequerimiento.RequerimientoPoliticaValMensaje(1, politica1);
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			resultado = frmRequerimiento.RequerimientoPoliticaVal(2,(int)this.cmbEstado.Value, idRequerimiento, (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			resultado = frmRequerimiento.RequerimientoPoliticaVal(3,(int)this.cmbEstado.Value, idRequerimiento, (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			resultado = frmRequerimiento.RequerimientoPoliticaVal(5,(int)this.cmbEstado.Value, idRequerimiento, (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			resultado = frmRequerimiento.RequerimientoPoliticaVal(6,(int)this.cmbEstado.Value, idRequerimiento, (int)this.cmbPersonal.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validar politicas
			
			#region	Generar rutas del documento
				
			string sCarpeta = "";
			string sUrlNube = "";
			string sRuta = "";
			string carpeta = "";
			string carpeta_3 = "";
			string carpeta_1 = "";
			string carpeta_2 = "";
			string txtURL = "";
			string sextencion = "";
			int idRequerimientoUrgencia = 0;
			int idRequerimientoComplejidad = 0;
			string dtFechaEntregaAdm = "";
			string dtFechaEntregaTec = "";

			try{idRequerimientoUrgencia = (int)this.cmdRequerimientoUrgencia.Value;}catch(Exception ex){string sms = ex.Message;}
			try{idRequerimientoComplejidad = (int)this.cmdRequerimientoComplejidad.Value;}catch(Exception ex){string sms = ex.Message;}
			try{dtFechaEntregaAdm = Convert.ToDateTime(this.dtFechaEntregaAdm.Value).ToString("yyyyMMdd HH:mm:ss");}catch(Exception ex){string sms = ex.Message;}
			try{dtFechaEntregaTec = Convert.ToDateTime(this.dtFechaEntregaTec.Value).ToString("yyyyMMdd HH:mm:ss");}catch(Exception ex){string sms = ex.Message;}

			carpeta_1 = @"\\192.168.1.10\Proyectos\RequerimientoDocumento";
			carpeta_2 = @"\\192.168.1.10\Proyectos\RequerimientoDocumento\"+MenuLatinium.stDirFacturacion;
				
			using (SqlConnection oConexion = new SqlConnection(frmRequerimiento.sconexionPoint))					
			{
				// Mostrar el modal de "procesando"
				MenuLatinium.MostrarModalProcesando();
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

					#region Garbar seguimiento
					//Guardar requerimiento para luego darle seguimiento
					int idSeguimiento = 0;
					string sSQLGraba = string.Format("Exec RequerimientoSeguimientoCrear '{0}',{1},{2},{3},{4},'{5}','{6}', {7},{8},'{9}','{10}', {11}", 
						(string)this.txtNota.Value, 
						(int)this.cmbPersonal.Value,
						idRequerimiento,
						(int)this.cmbEstado.Value,
						(int)this.cmbAsignado.Value,
						(string)this.txtNotaInterna.Value,
						this.cmbPersonal.Text,
						idRequerimientoUrgencia,
						idRequerimientoComplejidad,
						dtFechaEntregaAdm,
						dtFechaEntregaTec,
						frmRequerimiento.iRequerimientoControlMantenimientoLocales
						);

					oCmdActualiza.CommandText = sSQLGraba;
					idSeguimiento = (int)oCmdActualiza.ExecuteScalar();
					#endregion Grabar seguimiento

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugdDocumento.Rows.All)	
						if ((int)dr.Cells["idRequerimientoDocumento"].Value != 0)
						{
							#region Variables
							sCarpeta = "";
							sUrlNube = "";
							sRuta = "";
							carpeta = "";
							carpeta_3 = "";
							txtURL = "";
							sRuta = dr.Cells["txtDirectorioarchivo"].Value.ToString();		
							#endregion Variables
								
							#region carpetas
							carpeta = @"\\192.168.1.10\Proyectos\RequerimientoDocumento\"+MenuLatinium.stDirFacturacion+@"\" + FechaAsignacion + "_" + Convert.ToString (this.cmbResponsable.Text) + "_" + Numero;
							carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + FechaAsignacion + "_" + Convert.ToString (this.cmbResponsable.Text) + "_" +  Numero;
							#endregion carpetas

							#region directorio

							Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
							Directory.GetDirectories(@"\\192.168.1.10\Proyectos");

							if (!Directory.Exists(carpeta_1))
							{
								Directory.CreateDirectory(carpeta);
							}
							if (!Directory.Exists(carpeta_2))
							{
								Directory.CreateDirectory(carpeta);
							}
							if (!Directory.Exists(carpeta))
							{
								Directory.CreateDirectory(carpeta);
							}
							#endregion directorio
										
							#region Generar ruta documento
							txtDirectorioarchivo = sRuta;
							txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
							//Coger extensión del nombre del archivo para guardarlo con un nombre generado
							sextencion = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}'))) AS Archivo",txtDirectorioarchivo,"."));
							//@ARCHIVO_EXTENSION = RIGHT(@RUTA_ARCHIVO, CHARINDEX('.', REVERSE(@RUTA_ARCHIVO)))
							//Generar nombre documento
							string sSQLNumero = string.Format("Exec NumeracionLocales 89, 0, 0");
							oCmdActualiza.CommandText = sSQLNumero;
							string sNombreDocuento = oCmdActualiza.ExecuteScalar().ToString();
								
							carpeta  = (carpeta + @"\" + sNombreDocuento.Replace(" ", "_") + sextencion); // txtURL.Replace(" ", "_"));

										
							if (System.IO.File.Exists(carpeta))
							{
								MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								return;
							}
							else
							{
								System.IO.File.Copy (txtDirectorioarchivo ,carpeta, false);
								sCarpeta = carpeta;
								sUrlNube = carpeta_3.Replace(" ", "_")+ @"\" + sNombreDocuento.Replace(" ", "_") + sextencion;//+ txtURL.Replace(" ", "_");
								//MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
								//Si se guardo crear un registro por el nuevo guardado
								string sSQLGoogle = string.Format("Exec RequerimientoDocumentoNuevo '{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}',{7},{8},'{9}'", 
									MenuLatinium.stUsuario, 
									0, //Significa nuevo 
									(int)this.cmbResponsable.Value,
									(int)this.cmbPersonal.Value, 
									txtDirectorioarchivo,
									sCarpeta,
									sUrlNube,
									idSeguimiento,
									dr.Cells["idRequerimientoDocumento"].Value.ToString(),
									dtFecha.ToString("yyyyMMdd")
									);
								oCmdActualiza.CommandText = sSQLGoogle;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
					#endregion Generar ruta documento

					oTransaction.Commit();
					oConexion.Close();

					#region Enviar notificacion
					frmRequerimiento.RequerimientoNotificacionCrear(
						(int)this.cmbPersonal.Value, (int)this.cmbAsignado.Value, frmRequerimiento.iRequerimientoMensajeCambioEstado, idSeguimiento, (int)this.cmbEstado.Value, frmRequerimiento.iRequerimientoControlMantenimientoLocales);
					#endregion Enviar notificacion

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, registro guardado"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					//Funcion.WriteLog(ex.Message);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
				}
				finally
				{
					oConexion.Close();
					// Ocultar el modal de "procesando" una vez que se haya completado la consulta
					MenuLatinium.OcultarModalProcesando();
				}
			}

			#endregion Generar rutas del documento
		}

		private void ugdMenu_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//Propiedad que selecciona la pantalla del requerimiento
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.Header.Caption.ToString() == sColumna)
			{
				ActualizaValores((int)e.Cell.Row.Cells["id_seg"].Value,
					(bool)e.Cell.Row.Cells["Sel_ReqPantalla"].Value);
			}
		}

		private void ActualizaValores(int id_seg, bool Sel_ReqPantalla)
		{
			string sSQL = string.Format("Exec RequerimientoSelPantalla {0}, {1}", id_seg , Sel_ReqPantalla);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ugdDocumento_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			//e.Row.Cells["idRequerimientoDocumento"].Value = 0;
			//e.Row.Cells["UrlCompartido"].Value = "";
			//e.Row.Cells["idSeguimiento"].Value = 0;
			//e.Row.Cells["FechaDeposito"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); //DateTime.Today;
		}

		private void ugdDocumento_AfterRowUpdate(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		
		}

		private void ugdDocumento_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void ugdDocumento_ClickCellButton_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			bool Nuevo = true;  
			string txtURL ="";
			long itamano = 0;
			bool bResultado = false;
			string sMensaje = "";
			DataSet dsConfiguracion = new DataSet();

			#region Grabar seguimiento
			if (e.Cell.Column.ToString() == "btnDocumento")	
			{
				#region Validacion
				//if (!Validacion.vbTexto((string)this.dtFechaAsignacion.Value, "50", "50", "Fecha asignación")) return;
				//if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione un responsable")) return;
				#endregion Validacion
				#region Guardar ruta local
				if (Nuevo)
				{
					try
					{
						OpenFileDialog buscar = new OpenFileDialog ();
						if (buscar.ShowDialog () == DialogResult.OK)
						{

							//Sacar el tamaño del archivo subido
							txtDirectorioarchivo = buscar.FileName;
							System.IO.FileInfo info = new System.IO.FileInfo(txtDirectorioarchivo);
							//El lenght está en tamaños bytes
							itamano = info.Length;
						  
							dsConfiguracion = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfgVal {0}, {1}, {2}", frmRequerimiento.iConfiguracionDocumento, itamano, frmRequerimiento.iConfiguracionTamano));
							bResultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
							sMensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();
							
							if(bResultado)
							{
								txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
								bool bCont=Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec RequerimientoDocumentoNuevo '{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}',{7},{8}, '{9}'", 
									MenuLatinium.stUsuario, 
									1, //Significa nuevo 
									(int)this.cmbResponsable.Value,
									(int)this.cmbPersonal.Value, 
									txtDirectorioarchivo,
									"",
									"",
									0,
									0,
									dtFecha.ToString("yyyyMMdd")
									),false);

								if (bCont == false)
								{
									MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									return;
								}
							}
							else
							{
								MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}

							RequerimientoDocumentoCon(1);
						}	
					}
					catch (Exception ex)
					{
						MessageBox.Show ("Error: "+ ex.Message);
					}
				}
				else
				{
					#region error
					try
					{
						OpenFileDialog buscar = new OpenFileDialog ();
						if (buscar.ShowDialog () == DialogResult.OK)
						{
							txtDirectorioarchivo = buscar.FileName; 	
						}	
					}
					catch (Exception ex)
					{
						MessageBox.Show ("Error: "+ ex.Message);
					}
					#endregion error
				}
			}
			#endregion Guardar ruta local
		
			#endregion Grabar seguimiento
			#region Ver

			if (e.Cell.Column.ToString() == "btnVer")	
			{
				try
				{
					if(bVer)
					{
						sUrlNube = (string) e.Cell.Row.Cells["UrlNube"].Value;
						string UrlCompartido = (string) e.Cell.Row.Cells["UrlCompartido"].Value;
						string sCarpeta = (string) e.Cell.Row.Cells["sCarpeta"].Value;
						if (sUrlNube.Length > 10 )
						{
							Process.Start(sCarpeta);
						}
					}
					if(bNuevo)
					{
						txtDirectorioarchivo = (string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value;
						if (txtDirectorioarchivo.Length > 10 )
						{
							Process.Start(txtDirectorioarchivo);
						}
					}
				}
				catch (Exception ex)
				{
					string sms = ex.Message;
					return;						
				}
			}

			#endregion Ver
			#region Eliminar
			if(bNuevo)
			{
				if (e.Cell.Column.ToString() == "btnEliminar")	
				{
					try
					{
						int itamanodocumentos = this.ugdDocumento.Rows.Count;
						if(itamanodocumentos == 0)
						{
							return;
						}
						//Si se guardo crear un registro por el nuevo guardado
						bool bCont =Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec RequerimientoDocumentoEliminar '{0}', {1}", 
							MenuLatinium.stUsuario, 
							(int)e.Cell.Row.Cells["idRequerimientoDocumento"].Value //id del documento
							), false);
						if (bCont == true)
						{
							txtDirectorioarchivo = (string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value;
							txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
							// Eliminar el path
							//System.IO.File.Delete(txtDirectorioarchivo);	
							MessageBox.Show(string.Format("El archivo '{0}' se eliminó correctamente. ", txtURL.Replace(" ", "_")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							//return;
						}
						RequerimientoDocumentoCon(1);
					}
					catch (Exception ex)
					{
						string sms = ex.Message;
						return;						
					}
				}
			}
			#endregion Eliminar
		}

		private void RequerimientoDocumentoCon(int Nuevo)
		{
			this.ugdDocumento.DataSource = 
				FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoCon {0}, {1}, {2}, '{3}'", 
				Nuevo, (int)this.cmbResponsable.Value, (int)this.cmbPersonal.Value, Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd")
				)); 
		}

		private void RequerimientoDocumentoIdSeguimiento()
		{
			this.ugdDocumento.DataSource = 
				FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoIdSeguimiento {0}", 
				getidSeguimiento
				)); 
		}


		private void ugdDocumento_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if(bVer)
			{
				e.Layout.Bands[0].Columns["btndocumento"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				e.Layout.Bands[0].Columns["btnEliminar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
		}

		private void cmbPersonal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraDataSource1_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEstado_TextChanged(object sender, System.EventArgs e)
		{
			if(bNuevo)
			{
				string squery = "";
				try
				{
					//4 Mensaje default cerrado
					squery = String.Format("Exec RequerimientoMensajeRolCon {0},{1},{2}" , (int) this.cmbPersonal.Value, (int)this.cmbEstado.Value, 4);
					string sResultado = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
					if(sResultado != "")
					{
						this.txtNota.Text = sResultado;
					}
				}
				catch(Exception ex)
				{
					//Funcion.WriteLog(ex.Message);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				}
			}
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraDataSource2_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void label5_Click_1(object sender, System.EventArgs e)
		{
		
		}

		private void lblEstado_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter_1(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEstado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
