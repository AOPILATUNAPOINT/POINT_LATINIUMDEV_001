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
using Infragistics.Win.UltraWinGrid;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeguimientoAnadir.
	/// </summary>
	public class frmSeguimientoAnadir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.Label label10;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btSeguridad;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
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
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdMenu;
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
		public int idRequerimientoDepartamento = 0;
		public int idRequerimientoRol = 0;
		private System.Windows.Forms.TabPage tabPage6;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoAreaAtencionAsignaActivosObtener;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoAreaAtencionAsignaActivosObtener;
		private System.Windows.Forms.TabControl tabRequerimiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBodega;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDepartamento;
		private System.Windows.Forms.Label lblDepartamento;
		private System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcionCheckList;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCiudad;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCargo;
		public System.Windows.Forms.Label lblCargo;
		private System.Windows.Forms.Label lblCiudad;
		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Label lblDescripcion;
		private System.Windows.Forms.Label lblFechaTentativa;
		private System.Windows.Forms.Label lblFechaAsignacion;
		public string sRequerimientoRol = "";

		public frmSeguimientoAnadir(int IdResponsable,
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
		string numero,
		int idDepartamento)
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
			idRequerimientoDepartamento = idDepartamento;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoAreaAtencionAsigna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Requerimiento", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoAtencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoAreaAtencionAsigna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Requerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoAtencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id_seg");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombrePadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreHijo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_seg");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombrePadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imagen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimientoAnadir));
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 3);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imagen");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlCompartido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImg", 0);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlCompartido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblDescripcion = new System.Windows.Forms.Label();
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
			this.tabRequerimiento = new System.Windows.Forms.TabControl();
			this.tabPage6 = new System.Windows.Forms.TabPage();
			this.txtDescripcionCheckList = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.udgRequerimientoAreaAtencionAsignaActivosObtener = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoAreaAtencionAsignaActivosObtener = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.ugdMenu = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsMenu = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSeguridad = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.label10 = new System.Windows.Forms.Label();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.ugdSeguimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSeguimiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.txtNotaInterna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNotaInterna = new System.Windows.Forms.Label();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.imgEstado = new System.Windows.Forms.PictureBox();
			this.lblFechaTentativa = new System.Windows.Forms.Label();
			this.dtFechaTentativa = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblFechaAsignacion = new System.Windows.Forms.Label();
			this.dtFechaAsignacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
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
			this.txtCargo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCiudad = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCiudad = new System.Windows.Forms.Label();
			this.lblCargo = new System.Windows.Forms.Label();
			this.txtDepartamento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblDepartamento = new System.Windows.Forms.Label();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
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
			this.tabRequerimiento.SuspendLayout();
			this.tabPage6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionCheckList)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoAreaAtencionAsignaActivosObtener)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionAsignaActivosObtener)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).BeginInit();
			this.tabPage5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).BeginInit();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResponsable)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(326, 65);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(120, 26);
			this.label4.TabIndex = 13;
			this.label4.Text = "Empresa:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(432, 28);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(77, 26);
			this.label3.TabIndex = 11;
			this.label3.Text = "* Asignado:";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(10, 28);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(120, 26);
			this.label1.TabIndex = 9;
			this.label1.Text = "Responsable:";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(778, 28);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(67, 26);
			this.label2.TabIndex = 15;
			this.label2.Text = "Estado:";
			// 
			// lblTitulo
			// 
			this.lblTitulo.Location = new System.Drawing.Point(10, 18);
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.Size = new System.Drawing.Size(86, 37);
			this.lblTitulo.TabIndex = 16;
			this.lblTitulo.Text = "Título:";
			// 
			// lblDescripcion
			// 
			this.lblDescripcion.Location = new System.Drawing.Point(10, 65);
			this.lblDescripcion.Name = "lblDescripcion";
			this.lblDescripcion.Size = new System.Drawing.Size(96, 37);
			this.lblDescripcion.TabIndex = 18;
			this.lblDescripcion.Text = "Descripción:";
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
			this.txtDescripcion.Location = new System.Drawing.Point(106, 55);
			this.txtDescripcion.MaxLength = 2000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.ReadOnly = true;
			this.txtDescripcion.Size = new System.Drawing.Size(1197, 157);
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
			this.txtTitulo.Location = new System.Drawing.Point(106, 18);
			this.txtTitulo.MaxLength = 50;
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.ReadOnly = true;
			this.txtTitulo.Size = new System.Drawing.Size(1197, 22);
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
			this.cmbEstado.Location = new System.Drawing.Point(893, 28);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.ReadOnly = true;
			this.cmbEstado.Size = new System.Drawing.Size(134, 25);
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
			this.cmbAsignado.Location = new System.Drawing.Point(509, 28);
			this.cmbAsignado.Name = "cmbAsignado";
			this.cmbAsignado.Size = new System.Drawing.Size(249, 25);
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
			this.cmbEmpresas.Location = new System.Drawing.Point(432, 65);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(192, 25);
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
			this.cmbResponsable.Location = new System.Drawing.Point(346, 28);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(38, 25);
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
			// tabRequerimiento
			// 
			this.tabRequerimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabRequerimiento.Controls.Add(this.tabPage6);
			this.tabRequerimiento.Controls.Add(this.tabPage2);
			this.tabRequerimiento.Controls.Add(this.tabPage5);
			this.tabRequerimiento.Controls.Add(this.tabPage1);
			this.tabRequerimiento.Controls.Add(this.tabPage3);
			this.tabRequerimiento.Location = new System.Drawing.Point(19, 222);
			this.tabRequerimiento.Name = "tabRequerimiento";
			this.tabRequerimiento.SelectedIndex = 0;
			this.tabRequerimiento.Size = new System.Drawing.Size(1332, 371);
			this.tabRequerimiento.TabIndex = 806;
			// 
			// tabPage6
			// 
			this.tabPage6.Controls.Add(this.txtDescripcionCheckList);
			this.tabPage6.Controls.Add(this.label9);
			this.tabPage6.Controls.Add(this.udgRequerimientoAreaAtencionAsignaActivosObtener);
			this.tabPage6.Location = new System.Drawing.Point(4, 25);
			this.tabPage6.Name = "tabPage6";
			this.tabPage6.Size = new System.Drawing.Size(1324, 342);
			this.tabPage6.TabIndex = 4;
			this.tabPage6.Text = "Checklist";
			// 
			// txtDescripcionCheckList
			// 
			this.txtDescripcionCheckList.AcceptsReturn = true;
			this.txtDescripcionCheckList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcionCheckList.Appearance = appearance7;
			this.txtDescripcionCheckList.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDescripcionCheckList.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcionCheckList.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcionCheckList.Location = new System.Drawing.Point(19, 258);
			this.txtDescripcionCheckList.MaxLength = 2000;
			this.txtDescripcionCheckList.Multiline = true;
			this.txtDescripcionCheckList.Name = "txtDescripcionCheckList";
			this.txtDescripcionCheckList.ReadOnly = true;
			this.txtDescripcionCheckList.Size = new System.Drawing.Size(1576, 111);
			this.txtDescripcionCheckList.TabIndex = 815;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(19, 231);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(96, 26);
			this.label9.TabIndex = 814;
			this.label9.Text = "Descripción:";
			// 
			// udgRequerimientoAreaAtencionAsignaActivosObtener
			// 
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DataSource = this.udsRequerimientoAreaAtencionAsignaActivosObtener;
			appearance8.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Appearance = appearance8;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 280;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 653;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.Caption = "Área de atención";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 371;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn12.Header.Caption = "Asignar a";
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 531;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			ultraGridBand5.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.GroupByBox.Hidden = true;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Location = new System.Drawing.Point(19, 18);
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Name = "udgRequerimientoAreaAtencionAsignaActivosObtener";
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.Size = new System.Drawing.Size(1576, 153);
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.TabIndex = 813;
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgRequerimientoAreaAtencionAsignaActivosObtener_InitializeLayout);
			// 
			// udsRequerimientoAreaAtencionAsignaActivosObtener
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.udsRequerimientoAreaAtencionAsignaActivosObtener.Band.Columns.AddRange(new object[] {
																																																 ultraDataColumn9,
																																																 ultraDataColumn10,
																																																 ultraDataColumn11,
																																																 ultraDataColumn12});
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.ugdMenu);
			this.tabPage2.Controls.Add(this.btSeguridad);
			this.tabPage2.Controls.Add(this.btCancelar);
			this.tabPage2.Controls.Add(this.label10);
			this.tabPage2.Controls.Add(this.btAceptar);
			this.tabPage2.Location = new System.Drawing.Point(4, 25);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(1324, 342);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Módulo - Pantalla";
			// 
			// ugdMenu
			// 
			this.ugdMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdMenu.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdMenu.DataSource = this.udsMenu;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ugdMenu.DisplayLayout.Appearance = appearance13;
			this.ugdMenu.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 56;
			ultraGridColumn14.Header.Caption = "Seleccionar";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 141;
			ultraGridColumn15.AutoEdit = false;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.Caption = "Módulo";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 455;
			ultraGridColumn16.AutoEdit = false;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.Caption = "Pantalla";
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Width = 429;
			ultraGridColumn17.AutoEdit = false;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.Caption = "Subpantalla - Botón";
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Width = 472;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			ultraGridBand6.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdMenu.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ugdMenu.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdMenu.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdMenu.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdMenu.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ugdMenu.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 10F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdMenu.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdMenu.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdMenu.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ugdMenu.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdMenu.Location = new System.Drawing.Point(29, 18);
			this.ugdMenu.Name = "ugdMenu";
			this.ugdMenu.Size = new System.Drawing.Size(1250, 258);
			this.ugdMenu.TabIndex = 811;
			this.ugdMenu.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdMenu_AfterCellUpdate);
			// 
			// udsMenu
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			this.udsMenu.Band.Columns.AddRange(new object[] {
																												ultraDataColumn13,
																												ultraDataColumn14,
																												ultraDataColumn15,
																												ultraDataColumn16,
																												ultraDataColumn17});
			// 
			// btSeguridad
			// 
			this.btSeguridad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSeguridad.Location = new System.Drawing.Point(10, 433);
			this.btSeguridad.Name = "btSeguridad";
			this.btSeguridad.Size = new System.Drawing.Size(105, 27);
			this.btSeguridad.TabIndex = 14;
			this.btSeguridad.Text = "Con&figuracion";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(96, 470);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(90, 29);
			this.btCancelar.TabIndex = 13;
			this.btCancelar.Text = "&Cancelar";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(19, 9);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(55, 18);
			this.label10.TabIndex = 11;
			this.label10.Text = "* Grupo:";
			this.label10.Visible = false;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(10, 470);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(90, 29);
			this.btAceptar.TabIndex = 12;
			this.btAceptar.Text = "&Aceptar";
			// 
			// tabPage5
			// 
			this.tabPage5.Controls.Add(this.ugdSeguimiento);
			this.tabPage5.Location = new System.Drawing.Point(4, 25);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(1324, 342);
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
			appearance18.BackColor = System.Drawing.Color.White;
			this.ugdSeguimiento.DisplayLayout.Appearance = appearance18;
			ultraGridColumn18.AutoEdit = false;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn18.Header.Caption = "Seguimiento";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 79;
			ultraGridColumn19.AutoEdit = false;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Width = 136;
			ultraGridColumn20.AutoEdit = false;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 467;
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 75;
			ultraGridColumn22.Header.VisiblePosition = 9;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 120;
			ultraGridColumn23.Header.Caption = "Personal";
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Width = 227;
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 98;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Width = 231;
			ultraGridColumn26.Header.VisiblePosition = 11;
			ultraGridColumn26.Hidden = true;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			ultraGridColumn27.CellAppearance = appearance19;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			ultraGridColumn27.CellButtonAppearance = appearance20;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 19;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			ultraGridColumn28.CellAppearance = appearance21;
			appearance22.Image = ((object)(resources.GetObject("appearance22.Image")));
			ultraGridColumn28.CellButtonAppearance = appearance22;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 23;
			ultraGridColumn29.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn29.Header.Caption = "Estado";
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 123;
			ultraGridColumn30.AutoEdit = false;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn30.Header.Caption = "N°";
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Width = 26;
			ultraGridBand7.Columns.AddRange(new object[] {
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
			ultraGridBand7.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdSeguimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.ugdSeguimiento.DisplayLayout.GroupByBox.Hidden = true;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.ugdSeguimiento.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 10F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdSeguimiento.DisplayLayout.Override.HeaderAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSeguimiento.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSeguimiento.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.ugdSeguimiento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdSeguimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdSeguimiento.Location = new System.Drawing.Point(19, 18);
			this.ugdSeguimiento.Name = "ugdSeguimiento";
			this.ugdSeguimiento.Size = new System.Drawing.Size(1284, 313);
			this.ugdSeguimiento.TabIndex = 811;
			this.ugdSeguimiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ugdSeguimiento_InitializeRow);
			this.ugdSeguimiento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdSeguimiento_ClickCellButton);
			// 
			// udsSeguimiento
			// 
			ultraDataColumn18.DataType = typeof(System.UInt32);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.UInt32);
			ultraDataColumn22.DataType = typeof(System.UInt32);
			this.udsSeguimiento.Band.Columns.AddRange(new object[] {
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
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.txtNotaInterna);
			this.tabPage1.Controls.Add(this.lblNotaInterna);
			this.tabPage1.Controls.Add(this.lblTitulo);
			this.tabPage1.Controls.Add(this.txtTitulo);
			this.tabPage1.Controls.Add(this.txtDescripcion);
			this.tabPage1.Controls.Add(this.lblDescripcion);
			this.tabPage1.Location = new System.Drawing.Point(4, 25);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(1324, 342);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Descripción";
			this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
			// 
			// txtNotaInterna
			// 
			this.txtNotaInterna.AcceptsReturn = true;
			this.txtNotaInterna.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotaInterna.Appearance = appearance27;
			this.txtNotaInterna.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNotaInterna.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotaInterna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotaInterna.Location = new System.Drawing.Point(106, 231);
			this.txtNotaInterna.MaxLength = 2000;
			this.txtNotaInterna.Multiline = true;
			this.txtNotaInterna.Name = "txtNotaInterna";
			this.txtNotaInterna.ReadOnly = true;
			this.txtNotaInterna.Size = new System.Drawing.Size(1197, 37);
			this.txtNotaInterna.TabIndex = 1003;
			// 
			// lblNotaInterna
			// 
			this.lblNotaInterna.Location = new System.Drawing.Point(10, 240);
			this.lblNotaInterna.Name = "lblNotaInterna";
			this.lblNotaInterna.Size = new System.Drawing.Size(134, 27);
			this.lblNotaInterna.TabIndex = 1002;
			this.lblNotaInterna.Text = "Nota interna:";
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.ugdDocumento);
			this.tabPage3.Location = new System.Drawing.Point(4, 25);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(1324, 342);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Documento";
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance28.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance28;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 150;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn32.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn32.Header.VisiblePosition = 6;
			ultraGridColumn32.Width = 656;
			ultraGridColumn33.Header.Caption = "Seguimiento";
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 79;
			ultraGridColumn34.Header.Caption = "Dirección archivo";
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 400;
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 135;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 135;
			ultraGridColumn37.Header.Caption = "Dirección documento";
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 493;
			ultraGridColumn38.Header.VisiblePosition = 9;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 154;
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 73;
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 96;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn41.Header.VisiblePosition = 5;
			ultraGridColumn41.Width = 692;
			appearance29.Image = ((object)(resources.GetObject("appearance29.Image")));
			ultraGridColumn42.CellAppearance = appearance29;
			appearance30.Image = ((object)(resources.GetObject("appearance30.Image")));
			ultraGridColumn42.CellButtonAppearance = appearance30;
			ultraGridColumn42.Header.Caption = "...";
			ultraGridColumn42.Header.VisiblePosition = 4;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn42.Width = 42;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn42});
			ultraGridBand8.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.ugdDocumento.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 10F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.ugdDocumento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(10, 18);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(1284, 295);
			this.ugdDocumento.TabIndex = 810;
			this.ugdDocumento.AfterRowUpdate += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowUpdate);
			this.ugdDocumento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugdDocumento_AfterRowInsert);
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			this.ugdDocumento.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_AfterCellUpdate);
			this.ugdDocumento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdDocumento_InitializeLayout);
			// 
			// udsDocumento
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn27,
																														 ultraDataColumn28,
																														 ultraDataColumn29,
																														 ultraDataColumn30,
																														 ultraDataColumn31,
																														 ultraDataColumn32,
																														 ultraDataColumn33,
																														 ultraDataColumn34,
																														 ultraDataColumn35,
																														 ultraDataColumn36,
																														 ultraDataColumn37});
			// 
			// imgEstado
			// 
			this.imgEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.imgEstado.Location = new System.Drawing.Point(845, 28);
			this.imgEstado.Name = "imgEstado";
			this.imgEstado.Size = new System.Drawing.Size(235, 27);
			this.imgEstado.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.imgEstado.TabIndex = 1005;
			this.imgEstado.TabStop = false;
			// 
			// lblFechaTentativa
			// 
			this.lblFechaTentativa.Location = new System.Drawing.Point(1144, 32);
			this.lblFechaTentativa.Name = "lblFechaTentativa";
			this.lblFechaTentativa.Size = new System.Drawing.Size(48, 27);
			this.lblFechaTentativa.TabIndex = 1001;
			this.lblFechaTentativa.Text = "Fecha tentativa:";
			this.lblFechaTentativa.Visible = false;
			this.lblFechaTentativa.Click += new System.EventHandler(this.label8_Click);
			// 
			// dtFechaTentativa
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaTentativa.Appearance = appearance35;
			this.dtFechaTentativa.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaTentativa.DateButtons.Add(dateButton1);
			this.dtFechaTentativa.Enabled = false;
			this.dtFechaTentativa.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaTentativa.Location = new System.Drawing.Point(1240, 32);
			this.dtFechaTentativa.Name = "dtFechaTentativa";
			this.dtFechaTentativa.NonAutoSizeHeight = 23;
			this.dtFechaTentativa.ReadOnly = true;
			this.dtFechaTentativa.Size = new System.Drawing.Size(56, 24);
			this.dtFechaTentativa.SpinButtonsVisible = true;
			this.dtFechaTentativa.TabIndex = 1000;
			this.dtFechaTentativa.Value = ((object)(resources.GetObject("dtFechaTentativa.Value")));
			this.dtFechaTentativa.Visible = false;
			// 
			// lblFechaAsignacion
			// 
			this.lblFechaAsignacion.Location = new System.Drawing.Point(10, 102);
			this.lblFechaAsignacion.Name = "lblFechaAsignacion";
			this.lblFechaAsignacion.Size = new System.Drawing.Size(124, 36);
			this.lblFechaAsignacion.TabIndex = 999;
			this.lblFechaAsignacion.Text = "Fecha asignación:";
			// 
			// dtFechaAsignacion
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAsignacion.Appearance = appearance36;
			this.dtFechaAsignacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAsignacion.DateButtons.Add(dateButton2);
			this.dtFechaAsignacion.Enabled = false;
			this.dtFechaAsignacion.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFechaAsignacion.Location = new System.Drawing.Point(125, 102);
			this.dtFechaAsignacion.Name = "dtFechaAsignacion";
			this.dtFechaAsignacion.NonAutoSizeHeight = 23;
			this.dtFechaAsignacion.ReadOnly = true;
			this.dtFechaAsignacion.Size = new System.Drawing.Size(192, 24);
			this.dtFechaAsignacion.SpinButtonsVisible = true;
			this.dtFechaAsignacion.TabIndex = 998;
			this.dtFechaAsignacion.Value = ((object)(resources.GetObject("dtFechaAsignacion.Value")));
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
			this.btnSalir.Location = new System.Drawing.Point(1255, 65);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 26);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbTipoRequerimiento
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRequerimiento.Appearance = appearance37;
			this.cmbTipoRequerimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoRequerimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRequerimiento.DataSource = this.udsRequerimientoTipo;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn43,
																										 ultraGridColumn44});
			this.cmbTipoRequerimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbTipoRequerimiento.DisplayMember = "Descripcion";
			this.cmbTipoRequerimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRequerimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRequerimiento.Enabled = false;
			this.cmbTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRequerimiento.Location = new System.Drawing.Point(778, 65);
			this.cmbTipoRequerimiento.Name = "cmbTipoRequerimiento";
			this.cmbTipoRequerimiento.Size = new System.Drawing.Size(302, 25);
			this.cmbTipoRequerimiento.TabIndex = 809;
			this.cmbTipoRequerimiento.ValueMember = "idRequerimientoTipo";
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn38.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn38,
																																		 ultraDataColumn39});
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(672, 65);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(120, 26);
			this.label11.TabIndex = 808;
			this.label11.Text = "Requerimiento:";
			this.label11.Click += new System.EventHandler(this.label11_Click);
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
			this.btnNuevo2.Location = new System.Drawing.Point(1255, 18);
			this.btnNuevo2.Name = "btnNuevo2";
			this.btnNuevo2.Size = new System.Drawing.Size(96, 28);
			this.btnNuevo2.TabIndex = 1021;
			this.btnNuevo2.Text = "&Nuevo";
			this.btnNuevo2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo2, "Nuevo seguimiento");
			this.btnNuevo2.Click += new System.EventHandler(this.button1_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(10, 65);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(52, 18);
			this.label12.TabIndex = 1035;
			this.label12.Text = "Numero";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance38;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(125, 65);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(192, 25);
			this.txtNumero.TabIndex = 1034;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-826, 194);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(3041, 9);
			this.groupBox1.TabIndex = 1036;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.txtCargo);
			this.groupBox2.Controls.Add(this.txtCiudad);
			this.groupBox2.Controls.Add(this.lblCiudad);
			this.groupBox2.Controls.Add(this.lblCargo);
			this.groupBox2.Controls.Add(this.txtDepartamento);
			this.groupBox2.Controls.Add(this.lblDepartamento);
			this.groupBox2.Controls.Add(this.txtBodega);
			this.groupBox2.Controls.Add(this.lblBodega);
			this.groupBox2.Controls.Add(this.txtEmpresa);
			this.groupBox2.Controls.Add(this.txtResponsable);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.cmbResponsable);
			this.groupBox2.Controls.Add(this.cmbTipoRequerimiento);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.cmbAsignado);
			this.groupBox2.Controls.Add(this.cmbEmpresas);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.dtFechaTentativa);
			this.groupBox2.Controls.Add(this.dtFechaAsignacion);
			this.groupBox2.Controls.Add(this.lblFechaTentativa);
			this.groupBox2.Controls.Add(this.imgEstado);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.cmbEstado);
			this.groupBox2.Controls.Add(this.lblFechaAsignacion);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Location = new System.Drawing.Point(19, 9);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1217, 176);
			this.groupBox2.TabIndex = 1037;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// txtCargo
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCargo.Appearance = appearance39;
			this.txtCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCargo.Enabled = false;
			this.txtCargo.Location = new System.Drawing.Point(128, 138);
			this.txtCargo.Name = "txtCargo";
			this.txtCargo.Size = new System.Drawing.Size(278, 24);
			this.txtCargo.TabIndex = 1076;
			// 
			// txtCiudad
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCiudad.Appearance = appearance40;
			this.txtCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCiudad.Enabled = false;
			this.txtCiudad.Location = new System.Drawing.Point(504, 136);
			this.txtCiudad.Name = "txtCiudad";
			this.txtCiudad.Size = new System.Drawing.Size(259, 24);
			this.txtCiudad.TabIndex = 1075;
			// 
			// lblCiudad
			// 
			this.lblCiudad.AutoSize = true;
			this.lblCiudad.BackColor = System.Drawing.Color.Transparent;
			this.lblCiudad.Location = new System.Drawing.Point(432, 136);
			this.lblCiudad.Name = "lblCiudad";
			this.lblCiudad.Size = new System.Drawing.Size(51, 18);
			this.lblCiudad.TabIndex = 1074;
			this.lblCiudad.Text = "Ciudad:";
			this.lblCiudad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCargo
			// 
			this.lblCargo.Location = new System.Drawing.Point(16, 138);
			this.lblCargo.Name = "lblCargo";
			this.lblCargo.Size = new System.Drawing.Size(120, 27);
			this.lblCargo.TabIndex = 1072;
			this.lblCargo.Text = "Cargo solicitado:";
			// 
			// txtDepartamento
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepartamento.Appearance = appearance41;
			this.txtDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepartamento.Enabled = false;
			this.txtDepartamento.Location = new System.Drawing.Point(816, 96);
			this.txtDepartamento.Name = "txtDepartamento";
			this.txtDepartamento.Size = new System.Drawing.Size(264, 24);
			this.txtDepartamento.TabIndex = 1041;
			this.txtDepartamento.Visible = false;
			// 
			// lblDepartamento
			// 
			this.lblDepartamento.Location = new System.Drawing.Point(712, 96);
			this.lblDepartamento.Name = "lblDepartamento";
			this.lblDepartamento.Size = new System.Drawing.Size(105, 27);
			this.lblDepartamento.TabIndex = 1040;
			this.lblDepartamento.Text = "Departamento:";
			this.lblDepartamento.Visible = false;
			// 
			// txtBodega
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBodega.Appearance = appearance42;
			this.txtBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodega.Enabled = false;
			this.txtBodega.Location = new System.Drawing.Point(432, 104);
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.Size = new System.Drawing.Size(259, 24);
			this.txtBodega.TabIndex = 1039;
			// 
			// lblBodega
			// 
			this.lblBodega.Location = new System.Drawing.Point(328, 104);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 26);
			this.lblBodega.TabIndex = 1038;
			this.lblBodega.Text = "Local:";
			// 
			// txtEmpresa
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpresa.Appearance = appearance43;
			this.txtEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresa.Enabled = false;
			this.txtEmpresa.Location = new System.Drawing.Point(432, 65);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(230, 24);
			this.txtEmpresa.TabIndex = 1037;
			// 
			// txtResponsable
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResponsable.Appearance = appearance44;
			this.txtResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResponsable.Enabled = false;
			this.txtResponsable.Location = new System.Drawing.Point(125, 28);
			this.txtResponsable.Name = "txtResponsable";
			this.txtResponsable.Size = new System.Drawing.Size(288, 24);
			this.txtResponsable.TabIndex = 1036;
			// 
			// frmSeguimientoAnadir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1370, 618);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo2);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.tabRequerimiento);
			this.Name = "frmSeguimientoAnadir";
			this.Text = "Añadir seguimiento";
			this.Load += new System.EventHandler(this.frmSeguimientoAnadir_Load);
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
			this.tabRequerimiento.ResumeLayout(false);
			this.tabPage6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionCheckList)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoAreaAtencionAsignaActivosObtener)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionAsignaActivosObtener)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).EndInit();
			this.tabPage5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSeguimiento)).EndInit();
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotaInterna)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTentativa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
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
		string txtDirectorioarchivo ="";


		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmSeguimientoAnadir_Load(object sender, System.EventArgs e)
		{

			//Ubicar en que departamento se encuentran
			string sidRequerimientoDepartamento = "";
			if(idRequerimientoDepartamento > 0)
			{
				sidRequerimientoDepartamento = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoDepartamentoModel '{0}',{1}","Descripcion",idRequerimientoDepartamento));
			}
			if(sidRequerimientoDepartamento != "")
			{
				this.Text += " (" + sidRequerimientoDepartamento+ ")";
			}

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
			//Datos del usuario como su rol
			RequerimientoRolUsuario();
			//CargarAsignado
			this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}, {3}", idResponsable, idAsignado, AsignaTecnico, idRequerimientoDepartamento /*Aun no se asigna*/));
			this.cmbAsignado.Value = idAsignado;
			//Cargar EstadoRequerimiento
			this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoEstado {0}", idRequerimientoDepartamento));
			this.cmbEstado.Value = idEstado;
			//Cargar tipo de requermiento
			this.cmbTipoRequerimiento.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoTipoConsultar {0}, {1}", idRequerimientoDepartamento, idResponsable));
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
			this.txtDescripcionCheckList.Value = Descripcion;
			//Cargar NotaInterna
			this.txtNotaInterna.Value = NotaInterna;
			//Cargar el menú seguridad, pero solo de los accesos que tiene cada grupo asignado.
			RequerimientoMenuIdPantalla();
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
			//Ocultar panel seguridad si no es del departamento de sistemas
			OcultarPestañaPanel();
			RequerimientoAreaAtencionAsignaActivosObtener();
			//Cargar datos Requerimiento.
			RequerimientoModel();
			ConfiguracionPorDepartamento();
		}

		private void ConfiguracionPorDepartamento()
		{
			if(idRequerimientoDepartamento == frmRequerimiento.iRequerimientoDePersonal)
			{
				this.lblDepartamento.Text = "Área solicitante:";
				this.lblTitulo.Text = "Motivo vacante:";
				this.lblDescripcion.Text = "Observaciones adicionales:";
				this.lblFechaTentativa.Text = "Fecha estimada:";
				this.lblFechaAsignacion.Text = "Fecha del requerimiento:";
				this.lblBodega.Text = "Sucursal:";
			}
		}

		//Consulat model de requerimiento
		private void RequerimientoModel()
		{
			//Locales

			try
			{
				string squery = String.Format("Exec RequerimientoModel '{0}', {1}", "NombreBodega", idRequerimiento);
				this.txtBodega.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				if(this.txtBodega.Text == "")
				{
					this.txtBodega.Visible = false;
					this.lblBodega.Visible = false;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtBodega.Visible = false;
				this.lblBodega.Visible = false;
			}

			try
			{
				string squery = String.Format("Exec RequerimientoModel '{0}', {1}", "NombreCargo", idRequerimiento);
				this.txtCargo.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				if(this.txtCargo.Text == "")
				{
					this.txtCargo.Visible = false;
					this.lblCargo.Visible = false;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtCargo.Visible = false;
				this.lblCargo.Visible = false;
			}

			try
			{
				string squery = String.Format("Exec RequerimientoModel '{0}', {1}", "NombreCiudad", idRequerimiento);
				this.txtCiudad.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				if(this.txtCiudad.Text == "")
				{
					this.txtCiudad.Visible = false;
					this.lblCiudad.Visible = false;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtCiudad.Visible = false;
				this.lblCiudad.Visible = false;
			}


			//Departamento
			if(frmSeguimiento.RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idRequerimientoTipo, "bDepartamento"))
			{
				this.txtDepartamento.Visible = true;
				this.lblDepartamento.Visible = true;
				try
				{
					string squery = String.Format("Exec RequerimientoModel '{0}', {1}", "NombreDepartamento", idRequerimiento);
					this.txtDepartamento.Text = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
					if(this.txtDepartamento.Text == "")
					{
						this.txtDepartamento.Visible = false;
						this.lblDepartamento.Visible = false;
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					this.txtDepartamento.Visible = false;
					this.lblDepartamento.Visible = false;
				}
			}
		}

//		private void RequerimientoRolUsuario()
//		{
//			//Cual es el rol del usuario logeado
//			idRequerimientoRol = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento), false);
//			string squery = String.Format("Exec RequerimientoRolModel '{0}', {1}", "Subtitulo", idRequerimientoRol);
//			sRequerimientoRol = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
//		}

		private void RequerimientoAreaAtencionAsignaActivosObtener()
		{
			string squery = String.Format("Exec RequerimientoAreaAtencionAsignaActivosObtener {0}", idRequerimiento);
			this.udgRequerimientoAreaAtencionAsignaActivosObtener.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoRolUsuario()
		{
			//Cual es el rol del usuario logeado
			idRequerimientoRol = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento), false);
			string squery = String.Format("Exec RequerimientoRolModel '{0}', {1}", "Subtitulo", idRequerimientoRol);
			sRequerimientoRol = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
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
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0},'{1}', {2}", 
					idPersonalLogin,
					frmRequerimiento.iSeguimiento,
					idRequerimientoDepartamento
					), false); 
				//Permiso Valores técnicos
				bool bvalorestecnicos =
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0},'{1}', {3}", 
					idPersonalLogin,
					frmRequerimiento.iValoresTecnicos,
					idRequerimientoDepartamento
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

		private void RequerimientoMenuIdPantalla()
		{
			try
			{
				// Ejecutar SP y cargar DataTable
				DataTable dt = FuncionesProcedimientos.dtEjecucionGeneral(
					frmRequerimiento.sconexionPoint, 
					string.Format("EXEC RequerimientoMenuIdPantalla {0}", idRequerimiento)
					);
				// Asignar al grid
				this.ugdMenu.DataSource = dt;
				// Aplicar ordenamiento
				this.ugdMenu.DisplayLayout.Bands[0].SortedColumns.Add("NombrePadre", false, true);
				this.ugdMenu.DisplayLayout.Bands[0].SortedColumns.Add("NombreHijo", false, true);
				this.ugdMenu.DisplayLayout.Bands[0].SortedColumns.Add("NombreSubHijo", false, true);
				// Contar filas
				int totalFilas = 0;
				if (dt != null)
					totalFilas = dt.Rows.Count;
				// Actualizar el tabPage2
				tabPage2.Text = "Módulo - Pantalla (" + totalFilas + ")";
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

//		private void ControlCerrado()
//		{
//			if(idEstado == idCerrado)
//			{
//				this.btnNuevo2.Enabled = false;
//			}
//			else
//			{
//				this.btnNuevo2.Enabled = true;
//			}
//		}

		private void RequerimientoDocumentoIdSeguimiento()
		{
			try
			{
				// Ejecutar procedimiento y obtener DataTable
				DataTable dt = FuncionesProcedimientos.dtEjecucionGeneral(
					frmRequerimiento.sconexionPoint,
					string.Format("EXEC RequerimientoDocumentoIdSeguimiento {0}", idSeguimiento)
					);

				// Asignar al grid
				this.ugdDocumento.DataSource = dt;

				// Contar documentos
				int totalDocs = 0;
				if (dt != null)
					totalDocs = dt.Rows.Count;

				// Actualizar nombre del tab
				tabPage3.Text = "Documentos (" + totalDocs + ")";
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", 
					ex.GetType(), 
					ex.Message), 
					"Point Technology", 
					MessageBoxButtons.OK, 
					MessageBoxIcon.Error);
			}
		}

		private void RequerimientoSeguimientoCon()
		{
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();
			try
			{
				// Ejecutar consulta y asignar al grid
				DataTable dt = FuncionesProcedimientos.dtEjecucionGeneral(
					frmRequerimiento.sconexionPoint,
					string.Format("EXEC RequerimientoSeguimientoCon {0}", idRequerimiento)
					);
				this.ugdSeguimiento.DataSource = dt;
				// Contar filas
				int totalFilas = 0;
				if (dt != null)
					totalFilas = dt.Rows.Count;
				// Actualizar texto del tab
				// Lo dejo EXACTAMENTE como pediste: "Módulo - Pantalla (X)"
				tabPage5.Text = "Historial (" + totalFilas + ")";
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", 
					ex.GetType(), 
					ex.Message), 
					"Point Technology", 
					MessageBoxButtons.OK, 
					MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar modal
				MenuLatinium.OcultarModalProcesando();
			}
			// Mantengo tu funcionalidad existente
			this.ugdSeguimiento = Funcion.ContadorFilas(this.ugdSeguimiento, "Row");
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
					string UrlNube = (string) e.Cell.Row.Cells["UrlNube"].Value;
					if (UrlNube.Length > 10 )
					{
						Process.Start(UrlNube);
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
				using (frmSeguimientoAnadirNuevo frman = new frmSeguimientoAnadirNuevo(idRequerimiento, idResponsable, Numero, iNuevo, 0, idRequerimientoDepartamento))
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
						this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoAsignacionCon {0} , {1}, {2}, {3}", idResponsable, idAsignado, AsignaTecnico, idRequerimientoDepartamento /*Verificación de asignación*/));
						this.cmbAsignado.Value = idAsignado;
						this.txtNotaInterna.Text = NotaInterna;
						//Validar si lo cerro
						ControlCerrado();
						//Traer Imagen Estado
						ImagenEstado();
						//Cargar tipo de requermiento
						this.cmbTipoRequerimiento.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoTipoConsultar {0}, {1}", idRequerimientoDepartamento, idResponsable));
						this.cmbTipoRequerimiento.Value = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}', {1}", "idRequerimientoTipo", idRequerimiento),false);
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
						using (frmSeguimientoAnadirNuevo frman = new frmSeguimientoAnadirNuevo(idRequerimiento, idResponsable, Numero, iVer, getidSeguimiento, idRequerimientoDepartamento))
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
		private void OcultarPestañaPanel()
		{
			// Verifica si el idRequerimientoDepartamento es diferente del valor esperado
			if (!frmSeguimiento.RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idRequerimientoTipo,  "Pantalla"))
			{
				// Oculta tabPage2
				if (this.tabRequerimiento.TabPages.Contains(tabPage2))
				{
					tabRequerimiento.TabPages.Remove(tabPage2); // Remueve tabPage2 del TabControl
				}
			}
			else
			{
				// Asegúrate de que tabPage2 esté visible si la condición no se cumple
				if (!tabRequerimiento.TabPages.Contains(tabPage2))
				{
					tabRequerimiento.TabPages.Add(tabPage2); // Añade tabPage2 de nuevo al TabControl
				}
			}

			// Verifica si el idRequerimientoDepartamento es diferente del valor esperado
			if (!frmSeguimiento.RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idRequerimientoTipo, "Checklist"))
			{
				// Oculta tabPage4
				if (this.tabRequerimiento.TabPages.Contains(tabPage6))
				{
					tabRequerimiento.TabPages.Remove(tabPage6); // Remueve tabPage4 del TabControl
				}
			}
			else
			{
				// Asegúrate de que tabPage2 esté visible si la condición no se cumple
				if (!tabRequerimiento.TabPages.Contains(tabPage6))
				{
					tabRequerimiento.TabPages.Add(tabPage6); // Añade tabPage4 de nuevo al TabControl
				}
			}

			// Verifica si el idRequerimientoDepartamento es diferente del valor esperado
			if (!frmSeguimiento.RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idRequerimientoTipo,  "bDescripcion"))
			{
				// Oculta tabPage4
				if (this.tabRequerimiento.TabPages.Contains(tabPage1))
				{
					tabRequerimiento.TabPages.Remove(tabPage1); // Remueve tabPage4 del TabControl
				}
			}
			else
			{
				// Asegúrate de que tabPage2 esté visible si la condición no se cumple
				if (!tabRequerimiento.TabPages.Contains(tabPage1))
				{
					tabRequerimiento.TabPages.Add(tabPage1); // Añade tabPage4 de nuevo al TabControl
				}
			}
		}

		private void udgRequerimientoAreaAtencionAsignaActivosObtener_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			// Verificar el valor devuelto y ocultar la columna si es necesario
			if (!frmSeguimiento.RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idRequerimientoTipo,  "Asignacion"))
			{
				// Encuentra la columna por nombre
				UltraGridColumn column = e.Layout.Bands[0].Columns["Nombre"];
				if (column != null)
				{
					column.Hidden = true; // Oculta la columna
				}
			}
		}

		private void label8_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label11_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tabPage1_Click(object sender, System.EventArgs e)
		{
		
		}



	}
}
