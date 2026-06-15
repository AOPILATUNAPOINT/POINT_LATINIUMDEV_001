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
	/// Descripción breve de frmSeguimientoAnadirMantenimientoLocal.
	/// </summary>
	public class frmSeguimientoAnadirMantenimientoLocal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Button btnSalir;
		private System.ComponentModel.IContainer components;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTitulo;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaTentativa;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaAsignacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsMenu;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRequerimiento;
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.TabPage tabPage3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAsignado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.TabPage tabPage4;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSeguimiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.TabPage tabPage5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdSeguimiento;
		int idResponsable = 0;
		int idRequerimientoTipo = 0;
		int idEmpresa = 0;
		int idAsignado = 0;
		string Titulo = "";
		string Descripcion = "";
		DateTime FechaAsignacion;
		DateTime FechaTentativaEntrega;
		string NotaInterna = "";
		int idEstado = 0;
	  int idPantalla = 0;
		int idSeguimiento = 0;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		int idRequerimiento = 0;
		string Numero = "";
		int iNuevo = 1;
		int iEditar = 2;
		int iEliminar = 3;
		private System.Windows.Forms.GroupBox groupBox1;
		int iVer = 4;
		private System.Windows.Forms.Button btnNuevo2;
		int idCerrado = 7;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblNotaInterna;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotaInterna;
		int idPersonalLogin = 0;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtResponsable;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresa;
		bool AsignaTecnico = false;
		private System.Windows.Forms.PictureBox imgEstado;
		private System.Windows.Forms.ToolTip toolTip1;
		string Imagen = "";

		public frmSeguimientoAnadirMantenimientoLocal(int IdResponsable,
		int IdRequerimientoTipo,
		int IdEmpresa,
		int IdAsignado,
		string titulo,
		string descripcion,
		DateTime fechaAsignacion,
		DateTime fechaTentativaEntrega,
		string notaInterna,
		int IdEstado,
		int IdPantalla,
		int IdSeguimiento,
		int IdRequerimiento,
		string numero)
		{
			idResponsable = IdResponsable;
			idRequerimientoTipo = IdRequerimientoTipo;
			idEmpresa = IdEmpresa;
			idSeguimiento = IdSeguimiento;
			Titulo = titulo;
			Descripcion = descripcion;
			FechaAsignacion = fechaAsignacion;
			FechaTentativaEntrega = fechaTentativaEntrega;
			idPantalla = IdPantalla;
			idRequerimiento = IdRequerimiento;
			Numero = numero;
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idempresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idempresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimientoAnadirMantenimientoLocal));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imagen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 3);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imagen");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlCompartido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImg", 0);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlCompartido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_seg");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombrePadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbAsignado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.imgEstado = new System.Windows.Forms.PictureBox();
			this.txtNotaInterna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNotaInterna = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.dtFechaTentativa = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.dtFechaAsignacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.ugdSeguimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSeguimiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsMenu = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cmbTipoRequerimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label11 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnNuevo2 = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtResponsable = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).BeginInit();
			this.tabPage5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResponsable)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(272, 56);
			this.label4.Name = "label4";
			this.label4.TabIndex = 13;
			this.label4.Text = "* Empresa:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(424, 24);
			this.label3.Name = "label3";
			this.label3.TabIndex = 11;
			this.label3.Text = "* Asignado:";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 24);
			this.label1.Name = "label1";
			this.label1.TabIndex = 9;
			this.label1.Text = "* Responsable:";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(648, 160);
			this.label2.Name = "label2";
			this.label2.TabIndex = 15;
			this.label2.Text = "* Estado:";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(8, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(48, 23);
			this.label5.TabIndex = 16;
			this.label5.Text = "* Título:";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(8, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 23);
			this.label6.TabIndex = 18;
			this.label6.Text = "* Descripción:";
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(88, 48);
			this.txtDescripcion.MaxLength = 2000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.ReadOnly = true;
			this.txtDescripcion.Size = new System.Drawing.Size(840, 96);
			this.txtDescripcion.TabIndex = 799;
			// 
			// txtTitulo
			// 
			this.txtTitulo.AlwaysInEditMode = true;
			this.txtTitulo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTitulo.Appearance = appearance2;
			this.txtTitulo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTitulo.Location = new System.Drawing.Point(88, 16);
			this.txtTitulo.MaxLength = 50;
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.ReadOnly = true;
			this.txtTitulo.Size = new System.Drawing.Size(840, 19);
			this.txtTitulo.TabIndex = 800;
			// 
			// cmbEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance3;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(728, 160);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.ReadOnly = true;
			this.cmbEstado.Size = new System.Drawing.Size(200, 21);
			this.cmbEstado.TabIndex = 801;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// cmbAsignado
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAsignado.Appearance = appearance4;
			this.cmbAsignado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAsignado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignado.DataSource = this.udsAsignado;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbAsignado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAsignado.DisplayMember = "Nombre";
			this.cmbAsignado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAsignado.Enabled = false;
			this.cmbAsignado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignado.Location = new System.Drawing.Point(512, 24);
			this.cmbAsignado.Name = "cmbAsignado";
			this.cmbAsignado.Size = new System.Drawing.Size(328, 21);
			this.cmbAsignado.TabIndex = 802;
			this.cmbAsignado.ValueMember = "idPersonal";
			// 
			// udsAsignado
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsAsignado.Band.Columns.AddRange(new object[] {
																														ultraDataColumn3,
																														ultraDataColumn4});
			// 
			// cmbEmpresas
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance5;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresas.DataSource = this.udsEmpresa;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Enabled = false;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(344, 56);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(160, 21);
			this.cmbEmpresas.TabIndex = 803;
			this.cmbEmpresas.ValueMember = "idempresa";
			this.cmbEmpresas.Visible = false;
			// 
			// udsEmpresa
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsEmpresa.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn5,
																													 ultraDataColumn6});
			// 
			// cmbResponsable
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance6;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(384, 24);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(32, 21);
			this.cmbResponsable.TabIndex = 804;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.Visible = false;
			// 
			// udsResponsable
			// 
			ultraDataColumn7.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn7,
																															 ultraDataColumn8});
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage5);
			this.tabControl1.Controls.Add(this.tabPage3);
			this.tabControl1.Location = new System.Drawing.Point(16, 136);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(952, 296);
			this.tabControl1.TabIndex = 806;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.imgEstado);
			this.tabPage1.Controls.Add(this.txtNotaInterna);
			this.tabPage1.Controls.Add(this.lblNotaInterna);
			this.tabPage1.Controls.Add(this.label8);
			this.tabPage1.Controls.Add(this.dtFechaTentativa);
			this.tabPage1.Controls.Add(this.label7);
			this.tabPage1.Controls.Add(this.label5);
			this.tabPage1.Controls.Add(this.txtTitulo);
			this.tabPage1.Controls.Add(this.txtDescripcion);
			this.tabPage1.Controls.Add(this.label6);
			this.tabPage1.Controls.Add(this.dtFechaAsignacion);
			this.tabPage1.Controls.Add(this.cmbEstado);
			this.tabPage1.Controls.Add(this.label2);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(944, 270);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "* Descripción";
			// 
			// imgEstado
			// 
			this.imgEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.imgEstado.Location = new System.Drawing.Point(736, 160);
			this.imgEstado.Name = "imgEstado";
			this.imgEstado.Size = new System.Drawing.Size(192, 24);
			this.imgEstado.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.imgEstado.TabIndex = 1005;
			this.imgEstado.TabStop = false;
			// 
			// txtNotaInterna
			// 
			this.txtNotaInterna.AcceptsReturn = true;
			this.txtNotaInterna.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotaInterna.Appearance = appearance7;
			this.txtNotaInterna.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNotaInterna.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotaInterna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotaInterna.Location = new System.Drawing.Point(88, 200);
			this.txtNotaInterna.MaxLength = 2000;
			this.txtNotaInterna.Multiline = true;
			this.txtNotaInterna.Name = "txtNotaInterna";
			this.txtNotaInterna.ReadOnly = true;
			this.txtNotaInterna.Size = new System.Drawing.Size(840, 64);
			this.txtNotaInterna.TabIndex = 1003;
			// 
			// lblNotaInterna
			// 
			this.lblNotaInterna.Location = new System.Drawing.Point(8, 208);
			this.lblNotaInterna.Name = "lblNotaInterna";
			this.lblNotaInterna.Size = new System.Drawing.Size(112, 23);
			this.lblNotaInterna.TabIndex = 1002;
			this.lblNotaInterna.Text = "* Nota interna:";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(296, 160);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(120, 24);
			this.label8.TabIndex = 1001;
			this.label8.Text = "* Fecha tentativa:";
			// 
			// dtFechaTentativa
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaTentativa.Appearance = appearance8;
			this.dtFechaTentativa.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaTentativa.DateButtons.Add(dateButton1);
			this.dtFechaTentativa.Enabled = false;
			this.dtFechaTentativa.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaTentativa.Location = new System.Drawing.Point(416, 160);
			this.dtFechaTentativa.Name = "dtFechaTentativa";
			this.dtFechaTentativa.NonAutoSizeHeight = 23;
			this.dtFechaTentativa.ReadOnly = true;
			this.dtFechaTentativa.Size = new System.Drawing.Size(208, 21);
			this.dtFechaTentativa.SpinButtonsVisible = true;
			this.dtFechaTentativa.TabIndex = 1000;
			this.dtFechaTentativa.Value = ((object)(resources.GetObject("dtFechaTentativa.Value")));
			this.dtFechaTentativa.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaTentativa_BeforeDropDown);
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(8, 160);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 32);
			this.label7.TabIndex = 999;
			this.label7.Text = "* Fecha asignación:";
			// 
			// dtFechaAsignacion
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAsignacion.Appearance = appearance9;
			this.dtFechaAsignacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAsignacion.DateButtons.Add(dateButton2);
			this.dtFechaAsignacion.Enabled = false;
			this.dtFechaAsignacion.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaAsignacion.Location = new System.Drawing.Point(88, 160);
			this.dtFechaAsignacion.Name = "dtFechaAsignacion";
			this.dtFechaAsignacion.NonAutoSizeHeight = 23;
			this.dtFechaAsignacion.ReadOnly = true;
			this.dtFechaAsignacion.Size = new System.Drawing.Size(200, 21);
			this.dtFechaAsignacion.SpinButtonsVisible = true;
			this.dtFechaAsignacion.TabIndex = 998;
			this.dtFechaAsignacion.Value = ((object)(resources.GetObject("dtFechaAsignacion.Value")));
			// 
			// tabPage5
			// 
			this.tabPage5.Controls.Add(this.ugdSeguimiento);
			this.tabPage5.Location = new System.Drawing.Point(4, 22);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(944, 270);
			this.tabPage5.TabIndex = 3;
			this.tabPage5.Text = "Historial";
			// 
			// ugdSeguimiento
			// 
			this.ugdSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdSeguimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdSeguimiento.DataSource = this.udsSeguimiento;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ugdSeguimiento.DisplayLayout.Appearance = appearance10;
			ultraGridColumn9.AutoEdit = false;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Seguimiento";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 79;
			ultraGridColumn10.AutoEdit = false;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 136;
			ultraGridColumn11.AutoEdit = false;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 467;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.Header.Caption = "Personal";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 227;
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 98;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 231;
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Hidden = true;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			ultraGridColumn18.CellAppearance = appearance11;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			ultraGridColumn18.CellButtonAppearance = appearance12;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 19;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			ultraGridColumn19.CellAppearance = appearance13;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			ultraGridColumn19.CellButtonAppearance = appearance14;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 23;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.Header.Caption = "Estado";
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 123;
			ultraGridColumn21.AutoEdit = false;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.Caption = "N°";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 26;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn21});
			this.ugdSeguimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ugdSeguimiento.DisplayLayout.GroupByBox.Hidden = true;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ugdSeguimiento.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdSeguimiento.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSeguimiento.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSeguimiento.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ugdSeguimiento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdSeguimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdSeguimiento.Location = new System.Drawing.Point(32, 19);
			this.ugdSeguimiento.Name = "ugdSeguimiento";
			this.ugdSeguimiento.Size = new System.Drawing.Size(880, 205);
			this.ugdSeguimiento.TabIndex = 811;
			this.ugdSeguimiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ugdSeguimiento_InitializeRow);
			this.ugdSeguimiento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdSeguimiento_ClickCellButton);
			// 
			// udsSeguimiento
			// 
			ultraDataColumn9.DataType = typeof(System.UInt32);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(System.UInt32);
			ultraDataColumn13.DataType = typeof(System.UInt32);
			this.udsSeguimiento.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn9,
																															 ultraDataColumn10,
																															 ultraDataColumn11,
																															 ultraDataColumn12,
																															 ultraDataColumn13,
																															 ultraDataColumn14,
																															 ultraDataColumn15,
																															 ultraDataColumn16,
																															 ultraDataColumn17});
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.ugdDocumento);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(944, 270);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Documento *";
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance19.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance19;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 150;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn23.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Width = 421;
			ultraGridColumn24.Header.Caption = "Seguimiento";
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 79;
			ultraGridColumn25.Header.Caption = "Dirección archivo";
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 400;
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 135;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 135;
			ultraGridColumn28.Header.Caption = "Dirección documento";
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 493;
			ultraGridColumn29.Header.VisiblePosition = 9;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 154;
			ultraGridColumn30.Header.VisiblePosition = 10;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 73;
			ultraGridColumn31.Header.VisiblePosition = 11;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 96;
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Width = 445;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			ultraGridColumn33.CellAppearance = appearance20;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			ultraGridColumn33.CellButtonAppearance = appearance21;
			ultraGridColumn33.Header.Caption = "...";
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 25;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ugdDocumento.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ugdDocumento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(8, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(912, 232);
			this.ugdDocumento.TabIndex = 810;
			this.ugdDocumento.AfterRowUpdate += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowUpdate);
			this.ugdDocumento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowInsert);
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			this.ugdDocumento.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_AfterCellUpdate);
			this.ugdDocumento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdDocumento_InitializeLayout);
			// 
			// udsDocumento
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
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
																														 ultraDataColumn28});
			// 
			// udsMenu
			// 
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn29.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn30.DataType = typeof(bool);
			ultraDataColumn30.DefaultValue = false;
			this.udsMenu.Band.Columns.AddRange(new object[] {
																												ultraDataColumn29,
																												ultraDataColumn30,
																												ultraDataColumn31,
																												ultraDataColumn32,
																												ultraDataColumn33});
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
			this.btnSalir.Location = new System.Drawing.Point(896, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbTipoRequerimiento
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRequerimiento.Appearance = appearance26;
			this.cmbTipoRequerimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoRequerimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRequerimiento.DataSource = this.udsRequerimientoTipo;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbTipoRequerimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipoRequerimiento.DisplayMember = "Descripcion";
			this.cmbTipoRequerimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRequerimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRequerimiento.Enabled = false;
			this.cmbTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRequerimiento.Location = new System.Drawing.Point(616, 56);
			this.cmbTipoRequerimiento.Name = "cmbTipoRequerimiento";
			this.cmbTipoRequerimiento.Size = new System.Drawing.Size(224, 21);
			this.cmbTipoRequerimiento.TabIndex = 809;
			this.cmbTipoRequerimiento.ValueMember = "idRequerimientoTipo";
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn34.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn34,
																																		 ultraDataColumn35});
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(512, 56);
			this.label11.Name = "label11";
			this.label11.TabIndex = 808;
			this.label11.Text = "* Requerimiento:";
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
			// btnNuevo2
			// 
			this.btnNuevo2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo2.CausesValidation = false;
			this.btnNuevo2.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo2.Image")));
			this.btnNuevo2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo2.Location = new System.Drawing.Point(896, 16);
			this.btnNuevo2.Name = "btnNuevo2";
			this.btnNuevo2.Size = new System.Drawing.Size(80, 24);
			this.btnNuevo2.TabIndex = 1021;
			this.btnNuevo2.Text = "&Nuevo";
			this.btnNuevo2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo2, "Nuevo seguimiento");
			this.btnNuevo2.Click += new System.EventHandler(this.button1_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 56);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 16);
			this.label12.TabIndex = 1035;
			this.label12.Text = "Numero";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance27;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(88, 56);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(176, 21);
			this.txtNumero.TabIndex = 1034;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-688, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2392, 8);
			this.groupBox1.TabIndex = 1036;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.txtEmpresa);
			this.groupBox2.Controls.Add(this.txtResponsable);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.cmbResponsable);
			this.groupBox2.Controls.Add(this.cmbTipoRequerimiento);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.cmbAsignado);
			this.groupBox2.Controls.Add(this.cmbEmpresas);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(856, 100);
			this.groupBox2.TabIndex = 1037;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// txtEmpresa
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpresa.Appearance = appearance28;
			this.txtEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresa.Enabled = false;
			this.txtEmpresa.Location = new System.Drawing.Point(344, 56);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(160, 21);
			this.txtEmpresa.TabIndex = 1037;
			// 
			// txtResponsable
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResponsable.Appearance = appearance29;
			this.txtResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResponsable.Enabled = false;
			this.txtResponsable.Location = new System.Drawing.Point(104, 24);
			this.txtResponsable.Name = "txtResponsable";
			this.txtResponsable.Size = new System.Drawing.Size(312, 21);
			this.txtResponsable.TabIndex = 1036;
			// 
			// frmSeguimientoAnadirMantenimientoLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1000, 454);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo2);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.tabControl1);
			this.Name = "frmSeguimientoAnadirMantenimientoLocal";
			this.Text = "Añadir seguimiento";
			this.Load += new System.EventHandler(this.frmSeguimientoAnadirMantenimientoLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).EndInit();
			this.tabPage5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResponsable)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		bool bImpreso = false;
		private Acceso miAcceso;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;		
		#endregion Variables

		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";


		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmSeguimientoAnadirMantenimientoLocal_Load(object sender, System.EventArgs e)
		{
			idAsignado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT idPersonal FROM Seguimiento WHERE idSeguimiento = {0}", idSeguimiento), false);
			NotaInterna = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT NotaInterna FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento));
			idEstado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT idEstado FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento), false); //IdEstado;
			AsignaTecnico = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT AsignaTecnico FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento), false); //IdEstado;
			this.txtResponsable.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT Responsable FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento));
			this.txtEmpresa.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT Empresa FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento));
			//Cadena de conexión 192.168.1.165,56399
			SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'");
			SqlCommand oCmd = new SqlCommand();
			oCmd.Connection = oConexion;
			oCmd.CommandType = CommandType.Text;
			//CargarEmpresas
			oCmd.CommandText = string.Format("Select top 1 e.idEmpresa, e.Empresa From Empresas e Where e.idEmpresa = {0} Order By e.Empresa", idEmpresa);
			oConexion.Open();
			SqlDataAdapter oDA = new SqlDataAdapter(oCmd.CommandText, oConexion);
			DataSet oDS = new DataSet();
			oDA.Fill(oDS);
			this.cmbEmpresas.DataSource = oDS;
			this.cmbEmpresas.Value = idEmpresa;
			//CargarResponsable
			this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
			this.cmbResponsable.Value = idResponsable;
			//CargarAsignado
			this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}, {3}", idResponsable, idAsignado, AsignaTecnico /*Aun no se asigna*/, frmRequerimiento.iRequerimientoControlMantenimientoLocales));
			this.cmbAsignado.Value = idAsignado;
			//Cargar EstadoRequerimiento
			this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, "Exec RequerimientoEstado");
			this.cmbEstado.Value = idEstado;
			//Cargar tipo de requermiento
			this.cmbTipoRequerimiento.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoTipoConsultar {0}", frmRequerimiento.iRequerimientoControlMantenimientoLocales));
			this.cmbTipoRequerimiento.Value = idRequerimientoTipo;
			oConexion.Close();
			//Cargar la fecha de asignación
			this.dtFechaAsignacion.Value = FechaAsignacion;
			//Cargar la fecha tentativa
			this.dtFechaTentativa.Value = FechaTentativaEntrega;
			//Cargar Titulo
			this.txtTitulo.Value = Titulo;
			//Cargar Descripcion
			this.txtDescripcion.Value = Descripcion;
			//Cargar NotaInterna
			this.txtNotaInterna.Value = NotaInterna;
			//Cargar el menú seguridad, pero solo de los accesos que tiene cada grupo asignado.
			RequerimientoDocumentoIdSeguimiento();
			RequerimientoSeguimientoCon();
			//Cargar número de requerimiento
			this.txtNumero.Text = Numero;
			//Controles cuando el requerimiento se encuentra cerrado
			//ControlCerrado();
			//Carga id usuario tabla Inicio
			int idUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, String.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
			//Id personal de la que se logea
			idPersonalLogin = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("Exec RequerimientoResponsableConsultar {0}", idUsuarioInicio), false); //MenuLatinium.stUsuario //Producción se debe enviar el codigo variable global
			//Controles de seguimiento
			RequerimientoRolPermisoCon();
			//Bloquear botón nuevo cuando se encuentre cerrado
			ControlCerrado();
			//Traer Imagen Estado
			ImagenEstado();
		}

		private void ImagenEstado()
		{
			string squery = "";
			try
			{
				squery = String.Format("EXEC EstadoRequerimientoModel '{0}', {1}", "Imagen", idEstado);
				string Imagen = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				this.imgEstado.Image = Image.FromFile(Imagen); 
			}
			catch(Exception ex)
			{
				//Funcion.WriteLog(ex.Message);
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
			}
		}

		private void ControlCerrado()
		{
			try
			{
				if((int)this.cmbEstado.Value == frmRequerimiento.idCerrado)
				{
					this.btnNuevo2.Enabled = false;
				}
				//Bloquear botón nuevo si se rechazó de la misma manera, y podrá realizar en cualquier momento solo el adminsitrador.
				if((int)this.cmbEstado.Value == frmRequerimiento.iEstadoRequerimientoRechazado)
				{
					this.btnNuevo2.Enabled = false;
				}
			}
			catch(Exception ex)
			{
				//Funcion.WriteLog(ex.Message);
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
			}
			
		}

		private void RequerimientoRolPermisoCon()
		{
			try
			{
				//Permiso Seguimiento botón Nuevo
				this.btnNuevo2.Enabled = 
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0}, '{1}', {2}", 
					idPersonalLogin,
					frmRequerimiento.iSeguimiento,
					frmRequerimiento.iRequerimientoControlMantenimientoLocales
					), false); 
				//Permiso Valores técnicos
				bool bvalorestecnicos =
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0}, '{1}', {2}", 
					idPersonalLogin,
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

		private void RequerimientoDocumentoIdSeguimiento()
		{
			try
			{
				this.ugdDocumento.DataSource = 
					FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoIdSeguimiento {0}", 
					idSeguimiento
					)); 
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

		}

		private void RequerimientoSeguimientoCon()
		{
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();
			try
			{
				this.ugdSeguimiento.DataSource = 
					FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoSeguimientoCon {0}", 
					idRequerimiento
					)); 
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}
			this.ugdSeguimiento = Funcion.ContadorFilas(this.ugdSeguimiento, "Row");
		}

		private void RequerimientoDocumentoCon(int Nuevo)
		{
				this.ugdDocumento.DataSource = 
				FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoCon {0}, {1}, {2}", 
				Nuevo, (int)this.cmbResponsable.Value, (int)this.cmbAsignado.Value
				)); 
		}


		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
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

		private void ugdDocumento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{ 
			#region img
			try
			{
				if (e.Cell.Column.ToString() == "btnImg")
				{
					//				if (!miAcceso.BImagen) 
					//				{
					//					MessageBox.Show("No tiene acceso a Imagen ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					//					return;
					//				}

					sUrlNube = (string) e.Cell.Row.Cells["UrlNube"].Value;
					string UrlCompartido = (string) e.Cell.Row.Cells["UrlCompartido"].Value;
					string sCarpeta = (string) e.Cell.Row.Cells["sCarpeta"].Value;

					if (sUrlNube.Length > 10 )
					{
						Process.Start(sCarpeta);
						//System.Diagnostics.Process.Start("chrome.exe", sCarpeta);
					}
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			#endregion img
		}

		private void ugdDocumento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ugdDocumento_AfterRowUpdate(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		
		}

		private void ugdDocumento_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
				using (frmSeguimientoAnadirNuevoMantenimientoLocal frman = new frmSeguimientoAnadirNuevoMantenimientoLocal(idRequerimiento, idResponsable, Numero, iNuevo, 0))
				{		
					if (DialogResult.OK == frman.ShowDialog())
					{
					  RequerimientoSeguimientoCon();
						//Consultar de nuevo el estado del requerimiento
						idEstado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("Select idEstado from Requerimiento where idRequerimiento = {0}",idRequerimiento), false);
						this.cmbEstado.Value = idEstado;
						idAsignado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT idPersonal FROM Seguimiento WHERE idSeguimiento = {0}", idSeguimiento), false);
						NotaInterna = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT NotaInterna FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento));
						//CargarAsignado
						AsignaTecnico = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, String.Format("SELECT AsignaTecnico FROM Requerimiento WHERE idRequerimiento = {0}", idRequerimiento), false); //IdEstado;
						this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}", idResponsable, idAsignado, AsignaTecnico /*Verificación de asignación*/));
						this.cmbAsignado.Value = idAsignado;
						this.txtNotaInterna.Text = NotaInterna;
						//Validar si lo cerro
						ControlCerrado();
						//Traer Imagen Estado
						ImagenEstado();
					}
				}
		}

		private void ugdSeguimiento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{ 
				#region img
				if (e.Cell.Column.ToString() == "btnVer")
				{
					try
					{
						int getidSeguimiento = (int)e.Cell.Row.Cells["idSeguimiento"].Value;
						using (frmSeguimientoAnadirNuevoMantenimientoLocal frman = new frmSeguimientoAnadirNuevoMantenimientoLocal(idRequerimiento, idResponsable, Numero, iVer, getidSeguimiento))
						{		
							if (DialogResult.OK == frman.ShowDialog())
							{
					 
							}
						}
					}
					catch(Exception ex)
					{
						string sms = ex.Message;
						return;
					}
				}
				#endregion img
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ugdSeguimiento_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			try
			{
				e.Row.Cells["btnEstado"].ButtonAppearance.Image = Image.FromFile((string)e.Row.Cells["Imagen"].Value); 
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void dtFechaTentativa_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}



	}
}
