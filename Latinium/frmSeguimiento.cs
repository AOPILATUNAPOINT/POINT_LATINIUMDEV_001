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
using Infragistics.Win.UltraWinGrid; // Para manejar UltraGrid y eventos relacionados
using Latinium.Models;
using Latinium.Services;
using Latinium.Services.Requerimiento;
using Latinium.Models.Requerimiento;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeguimiento.
	/// </summary>
	public class frmSeguimiento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.Label label10;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btSeguridad;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.ComponentModel.IContainer components;
		int idResponsable = 0;
		int idAsignado = 0;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTitulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsMenu;
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
		int idEmpresa = 0;
		bool bVer = false;
		bool bNuevo = false;
		int iNuevo = 1;
		int iEditar = 2;
		int iEliminar = 3;
		int iVer = 4;
		int idRequerimiento = 0;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusquedaMenu;
		bool bRequerimientosDocumento = true;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotaIterna;
		private System.Windows.Forms.Label lblNotaInterna;
		private System.Windows.Forms.TabControl tabRequerimiento;
		string Title = "";
		public static string sconexionEmpresaSeguimiento = @"Password=Bl45o6$9;Persist Security Info=True;User ID= {0};Initial Catalog={1};Data Source={2}";
		string getsconexionEmpresaSeguimiento = "";
		int icontadoringreso = 0;
		private System.Windows.Forms.ToolTip toolTip1;
		int iTipo = 0;
		public int idRequerimientoDepartamento = 0;
		public int idRequerimientoRol = 0;
		public string sRequerimientoRol = "";
		private System.Windows.Forms.TabPage tabPage4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoAreaAtencionActivosObtener;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoAreaAtencionActivosObtener;
		private System.Windows.Forms.TabPage tabPage5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private System.Windows.Forms.Label label4;
		public System.Windows.Forms.Label label1;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRequerimientoDepartamento;
		private System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcionCheckList;
		private System.Windows.Forms.TabPage tabPage7;
		public System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Label lblDescripcion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargoSolicitado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSucursal;
		private System.Windows.Forms.TabControl tabDatosPrincipales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoCiudadExp;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoBodegaXIdCiudad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoCargoCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRequerimiento;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignado;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox grpAsignacion;
		//Ubicar en que departamento se encuentran
		public string sidRequerimientoDepartamento = "";
		private DocumentoService _documentoService;

		public frmSeguimiento(int IdResponsable, int IdAsignado, int IdEmpresa, int ITipo, int idDepartamento)
		{
			// Primero inicializa los componentes del formulario
			InitializeComponent();
			// Ahora asigna los valores a las variables locales/miembros
			idResponsable = IdResponsable;
			idAsignado = IdAsignado;
			idEmpresa = IdEmpresa;
			iTipo = ITipo;
			idRequerimientoDepartamento = idDepartamento;
			_documentoService = new DocumentoService(); // Inicializamos nuestro service

			// Aquí puedes agregar cualquier código adicional que necesites
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimiento));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoAreaAtencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AreaAtencionTipo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoAreaAtencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AreaAtencionTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id_seg");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombrePadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreHijo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id_seg");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel_ReqPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombrePadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreSubHijo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DocumentoRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 1);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 2);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DocumentoRef");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Id");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("name");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("name");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblDescripcion = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.tabRequerimiento = new System.Windows.Forms.TabControl();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDescripcionCheckList = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.ugdRequerimientoAreaAtencionActivosObtener = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoAreaAtencionActivosObtener = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.txtNotaIterna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNotaInterna = new System.Windows.Forms.Label();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.txtBusquedaMenu = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugdMenu = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsMenu = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSeguridad = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.label10 = new System.Windows.Forms.Label();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.spnGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.tabDatosPrincipales = new System.Windows.Forms.TabControl();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.grpAsignacion = new System.Windows.Forms.GroupBox();
			this.cmbAsignado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbTipoRequerimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbRequerimientoDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblDepartamento = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsProyecto = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage7 = new System.Windows.Forms.TabPage();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoBodegaXIdCiudad = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoCiudadExp = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbCargoSolicitado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoCargoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			this.tabRequerimiento.SuspendLayout();
			this.tabPage4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionCheckList)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoAreaAtencionActivosObtener)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionActivosObtener)).BeginInit();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaIterna)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBusquedaMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.tabDatosPrincipales.SuspendLayout();
			this.tabPage5.SuspendLayout();
			this.grpAsignacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRequerimientoDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsProyecto)).BeginInit();
			this.tabPage7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoBodegaXIdCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCiudadExp)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoSolicitado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCargoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(733, 62);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(60, 24);
			this.lblEstado.TabIndex = 15;
			this.lblEstado.Text = "* Estado:";
			this.lblEstado.Visible = false;
			// 
			// lblTitulo
			// 
			this.lblTitulo.Location = new System.Drawing.Point(8, 16);
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.Size = new System.Drawing.Size(72, 32);
			this.lblTitulo.TabIndex = 16;
			this.lblTitulo.Text = "* Título:";
			// 
			// lblDescripcion
			// 
			this.lblDescripcion.Location = new System.Drawing.Point(8, 56);
			this.lblDescripcion.Name = "lblDescripcion";
			this.lblDescripcion.Size = new System.Drawing.Size(80, 40);
			this.lblDescripcion.TabIndex = 18;
			this.lblDescripcion.Text = "* Descripción:";
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(88, 48);
			this.txtDescripcion.MaxLength = 2000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(1036, 104);
			this.txtDescripcion.TabIndex = 799;
			this.toolTip1.SetToolTip(this.txtDescripcion, "Descripción de lo que necesita que se realice para el requerimiento");
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtTitulo
			// 
			this.txtTitulo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTitulo.Appearance = appearance2;
			this.txtTitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTitulo.Location = new System.Drawing.Point(88, 16);
			this.txtTitulo.MaxLength = 50;
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.Size = new System.Drawing.Size(1036, 21);
			this.txtTitulo.TabIndex = 800;
			this.toolTip1.SetToolTip(this.txtTitulo, "Se decribe brevemente un título de su requerimiento.");
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
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(787, 62);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(246, 21);
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
			// udsResponsable
			// 
			ultraDataColumn7.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn7,
																															 ultraDataColumn8});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(1092, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 805;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// tabRequerimiento
			// 
			this.tabRequerimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabRequerimiento.Controls.Add(this.tabPage4);
			this.tabRequerimiento.Controls.Add(this.tabPage1);
			this.tabRequerimiento.Controls.Add(this.tabPage2);
			this.tabRequerimiento.Controls.Add(this.tabPage3);
			this.tabRequerimiento.Location = new System.Drawing.Point(16, 173);
			this.tabRequerimiento.Name = "tabRequerimiento";
			this.tabRequerimiento.SelectedIndex = 0;
			this.tabRequerimiento.Size = new System.Drawing.Size(1156, 437);
			this.tabRequerimiento.TabIndex = 806;
			// 
			// tabPage4
			// 
			this.tabPage4.Controls.Add(this.label9);
			this.tabPage4.Controls.Add(this.txtDescripcionCheckList);
			this.tabPage4.Controls.Add(this.btnNuevo);
			this.tabPage4.Controls.Add(this.ugdRequerimientoAreaAtencionActivosObtener);
			this.tabPage4.Location = new System.Drawing.Point(4, 22);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.Size = new System.Drawing.Size(1148, 411);
			this.tabPage4.TabIndex = 3;
			this.tabPage4.Text = "* Checklist";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(16, 208);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 1058;
			this.label9.Text = "* Descripción:";
			// 
			// txtDescripcionCheckList
			// 
			this.txtDescripcionCheckList.AcceptsReturn = true;
			this.txtDescripcionCheckList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcionCheckList.Appearance = appearance4;
			this.txtDescripcionCheckList.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcionCheckList.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcionCheckList.Location = new System.Drawing.Point(20, 229);
			this.txtDescripcionCheckList.MaxLength = 2000;
			this.txtDescripcionCheckList.Multiline = true;
			this.txtDescripcionCheckList.Name = "txtDescripcionCheckList";
			this.txtDescripcionCheckList.Size = new System.Drawing.Size(914, 111);
			this.txtDescripcionCheckList.TabIndex = 1057;
			this.toolTip1.SetToolTip(this.txtDescripcionCheckList, "Descripción de lo que necesita que se realice para el requerimiento");
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 381);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 25);
			this.btnNuevo.TabIndex = 1056;
			this.btnNuevo.Text = "&Nuevo item";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// ugdRequerimientoAreaAtencionActivosObtener
			// 
			this.ugdRequerimientoAreaAtencionActivosObtener.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoAreaAtencionActivosObtener.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoAreaAtencionActivosObtener.DataSource = this.udsRequerimientoAreaAtencionActivosObtener;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Appearance = appearance5;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 202;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.Caption = "Área de atención";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 248;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 110;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Asignar a";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 248;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 165;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Requerimiento";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 287;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoAreaAtencionActivosObtener.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdRequerimientoAreaAtencionActivosObtener.Location = new System.Drawing.Point(20, 21);
			this.ugdRequerimientoAreaAtencionActivosObtener.Name = "ugdRequerimientoAreaAtencionActivosObtener";
			this.ugdRequerimientoAreaAtencionActivosObtener.Size = new System.Drawing.Size(914, 184);
			this.ugdRequerimientoAreaAtencionActivosObtener.TabIndex = 812;
			this.ugdRequerimientoAreaAtencionActivosObtener.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoAreaAtencionActivosObtener_AfterCellUpdate);
			this.ugdRequerimientoAreaAtencionActivosObtener.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdRequerimientoAreaAtencionActivosObtener_InitializeLayout);
			// 
			// udsRequerimientoAreaAtencionActivosObtener
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(bool);
			this.udsRequerimientoAreaAtencionActivosObtener.Band.Columns.AddRange(new object[] {
																																													 ultraDataColumn9,
																																													 ultraDataColumn10,
																																													 ultraDataColumn11,
																																													 ultraDataColumn12,
																																													 ultraDataColumn13,
																																													 ultraDataColumn14});
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.txtNotaIterna);
			this.tabPage1.Controls.Add(this.lblNotaInterna);
			this.tabPage1.Controls.Add(this.lblTitulo);
			this.tabPage1.Controls.Add(this.txtTitulo);
			this.tabPage1.Controls.Add(this.txtDescripcion);
			this.tabPage1.Controls.Add(this.lblDescripcion);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(1148, 411);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "* Descripción";
			// 
			// txtNotaIterna
			// 
			this.txtNotaIterna.AcceptsReturn = true;
			this.txtNotaIterna.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotaIterna.Appearance = appearance10;
			this.txtNotaIterna.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotaIterna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotaIterna.Location = new System.Drawing.Point(88, 168);
			this.txtNotaIterna.MaxLength = 2000;
			this.txtNotaIterna.Multiline = true;
			this.txtNotaIterna.Name = "txtNotaIterna";
			this.txtNotaIterna.Size = new System.Drawing.Size(1036, 72);
			this.txtNotaIterna.TabIndex = 1004;
			// 
			// lblNotaInterna
			// 
			this.lblNotaInterna.Location = new System.Drawing.Point(8, 176);
			this.lblNotaInterna.Name = "lblNotaInterna";
			this.lblNotaInterna.Size = new System.Drawing.Size(112, 23);
			this.lblNotaInterna.TabIndex = 1002;
			this.lblNotaInterna.Text = "Nota interna:";
			this.lblNotaInterna.Click += new System.EventHandler(this.label9_Click);
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.btnConsultar);
			this.tabPage2.Controls.Add(this.txtBusquedaMenu);
			this.tabPage2.Controls.Add(this.ugdMenu);
			this.tabPage2.Controls.Add(this.btSeguridad);
			this.tabPage2.Controls.Add(this.btCancelar);
			this.tabPage2.Controls.Add(this.label10);
			this.tabPage2.Controls.Add(this.btAceptar);
			this.tabPage2.Controls.Add(this.spnGrupo);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(1148, 411);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = " Módulo - Pantalla";
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(256, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1020;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// txtBusquedaMenu
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusquedaMenu.Appearance = appearance11;
			this.txtBusquedaMenu.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusquedaMenu.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusquedaMenu.Location = new System.Drawing.Point(24, 8);
			this.txtBusquedaMenu.MaxLength = 200;
			this.txtBusquedaMenu.Name = "txtBusquedaMenu";
			this.txtBusquedaMenu.Size = new System.Drawing.Size(216, 21);
			this.txtBusquedaMenu.TabIndex = 1004;
			this.txtBusquedaMenu.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusquedaMenu_KeyPress);
			this.txtBusquedaMenu.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// ugdMenu
			// 
			this.ugdMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdMenu.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdMenu.DataSource = this.udsMenu;
			appearance12.BackColor = System.Drawing.Color.White;
			this.ugdMenu.DisplayLayout.Appearance = appearance12;
			this.ugdMenu.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 56;
			ultraGridColumn10.Header.Caption = "Seleccionar";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 270;
			ultraGridColumn11.Header.Caption = "Módulo";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 260;
			ultraGridColumn12.Header.Caption = "Pantalla";
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 254;
			ultraGridColumn13.Header.Caption = "Subpantalla - Botón";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 298;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdMenu.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ugdMenu.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdMenu.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdMenu.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdMenu.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ugdMenu.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdMenu.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdMenu.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdMenu.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ugdMenu.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdMenu.Location = new System.Drawing.Point(24, 40);
			this.ugdMenu.Name = "ugdMenu";
			this.ugdMenu.Size = new System.Drawing.Size(1103, 340);
			this.ugdMenu.TabIndex = 811;
			this.ugdMenu.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdMenu_ClickCellButton);
			this.ugdMenu.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdMenu_AfterCellUpdate);
			this.ugdMenu.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdMenu_InitializeLayout);
			// 
			// udsMenu
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			this.udsMenu.Band.Columns.AddRange(new object[] {
																												ultraDataColumn15,
																												ultraDataColumn16,
																												ultraDataColumn17,
																												ultraDataColumn18,
																												ultraDataColumn19});
			// 
			// btSeguridad
			// 
			this.btSeguridad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSeguridad.Location = new System.Drawing.Point(8, 491);
			this.btSeguridad.Name = "btSeguridad";
			this.btSeguridad.Size = new System.Drawing.Size(88, 25);
			this.btSeguridad.TabIndex = 14;
			this.btSeguridad.Text = "Con&figuracion";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(80, 523);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 26);
			this.btCancelar.TabIndex = 13;
			this.btCancelar.Text = "&Cancelar";
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(948, 8);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(46, 16);
			this.label10.TabIndex = 11;
			this.label10.Text = "* Grupo:";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 523);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 26);
			this.btAceptar.TabIndex = 12;
			this.btAceptar.Text = "&Aceptar";
			// 
			// spnGrupo
			// 
			this.spnGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnGrupo.Location = new System.Drawing.Point(1020, 8);
			this.spnGrupo.MaskInput = "nn";
			this.spnGrupo.MaxValue = 36;
			this.spnGrupo.MinValue = 1;
			this.spnGrupo.Name = "spnGrupo";
			this.spnGrupo.PromptChar = ' ';
			this.spnGrupo.ReadOnly = true;
			this.spnGrupo.Size = new System.Drawing.Size(104, 21);
			this.spnGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnGrupo.TabIndex = 16;
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.ugdDocumento);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(1148, 411);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Documento";
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance17.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance17;
			ultraGridColumn14.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 150;
			ultraGridColumn15.AutoEdit = false;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Format = "dd/MM/yyy HH:mm:ss";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 352;
			ultraGridColumn16.Header.Caption = "Seguimiento";
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 79;
			ultraGridColumn17.AutoEdit = false;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.Caption = "Dirección archivo";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 1218;
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 135;
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 135;
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 137;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 154;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 88;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Hidden = true;
			appearance18.Image = ((object)(resources.GetObject("appearance18.Image")));
			appearance18.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance18;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			appearance19.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellButtonAppearance = appearance19;
			ultraGridColumn24.Header.Caption = "...";
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 32;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			appearance20.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance20;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			appearance21.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellButtonAppearance = appearance21;
			ultraGridColumn25.Header.Caption = "...";
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 31;
			appearance22.Image = ((object)(resources.GetObject("appearance22.Image")));
			appearance22.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance22;
			appearance23.Image = ((object)(resources.GetObject("appearance23.Image")));
			appearance23.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellButtonAppearance = appearance23;
			ultraGridColumn26.Header.Caption = "...";
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 34;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn26});
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 10F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(16, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(1116, 192);
			this.ugdDocumento.TabIndex = 810;
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			// 
			// udsDocumento
			// 
			ultraDataColumn20.DataType = typeof(bool);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(object);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn20,
																														 ultraDataColumn21,
																														 ultraDataColumn22,
																														 ultraDataColumn23,
																														 ultraDataColumn24,
																														 ultraDataColumn25,
																														 ultraDataColumn26,
																														 ultraDataColumn27,
																														 ultraDataColumn28,
																														 ultraDataColumn29});
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
			this.btnSalir.Location = new System.Drawing.Point(1092, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn30,
																																		 ultraDataColumn31});
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-667, 153);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2580, 8);
			this.groupBox1.TabIndex = 1032;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// tabDatosPrincipales
			// 
			this.tabDatosPrincipales.Controls.Add(this.tabPage5);
			this.tabDatosPrincipales.Controls.Add(this.tabPage7);
			this.tabDatosPrincipales.Location = new System.Drawing.Point(16, 8);
			this.tabDatosPrincipales.Name = "tabDatosPrincipales";
			this.tabDatosPrincipales.SelectedIndex = 0;
			this.tabDatosPrincipales.Size = new System.Drawing.Size(1064, 138);
			this.tabDatosPrincipales.TabIndex = 1034;
			// 
			// tabPage5
			// 
			this.tabPage5.Controls.Add(this.grpAsignacion);
			this.tabPage5.Controls.Add(this.cmbRequerimientoDepartamento);
			this.tabPage5.Controls.Add(this.lblDepartamento);
			this.tabPage5.Controls.Add(this.lblBodega);
			this.tabPage5.Controls.Add(this.cmbResponsable);
			this.tabPage5.Controls.Add(this.cmbEmpresas);
			this.tabPage5.Controls.Add(this.label4);
			this.tabPage5.Controls.Add(this.label1);
			this.tabPage5.Controls.Add(this.cmbBodega);
			this.tabPage5.Controls.Add(this.cmbEstado);
			this.tabPage5.Controls.Add(this.lblEstado);
			this.tabPage5.Location = new System.Drawing.Point(4, 22);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(1056, 112);
			this.tabPage5.TabIndex = 0;
			this.tabPage5.Text = "Datos principales";
			// 
			// grpAsignacion
			// 
			this.grpAsignacion.Controls.Add(this.cmbAsignado);
			this.grpAsignacion.Controls.Add(this.label11);
			this.grpAsignacion.Controls.Add(this.cmbTipoRequerimiento);
			this.grpAsignacion.Controls.Add(this.label3);
			this.grpAsignacion.Location = new System.Drawing.Point(13, 42);
			this.grpAsignacion.Name = "grpAsignacion";
			this.grpAsignacion.Size = new System.Drawing.Size(714, 62);
			this.grpAsignacion.TabIndex = 1076;
			this.grpAsignacion.TabStop = false;
			this.grpAsignacion.Text = "Asignación";
			// 
			// cmbAsignado
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAsignado.Appearance = appearance28;
			this.cmbAsignado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAsignado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignado.DataSource = this.udsAsignado;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 305;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbAsignado.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbAsignado.DisplayMember = "Nombre";
			this.cmbAsignado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignado.Location = new System.Drawing.Point(100, 21);
			this.cmbAsignado.Name = "cmbAsignado";
			this.cmbAsignado.Size = new System.Drawing.Size(247, 21);
			this.cmbAsignado.TabIndex = 1073;
			this.cmbAsignado.ValueMember = "idPersonal";
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(380, 21);
			this.label11.Name = "label11";
			this.label11.TabIndex = 1074;
			this.label11.Text = "* Requerimiento:";
			// 
			// cmbTipoRequerimiento
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRequerimiento.Appearance = appearance29;
			this.cmbTipoRequerimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoRequerimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRequerimiento.DataSource = this.udsRequerimientoTipo;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.Caption = "Descripción";
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 520;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbTipoRequerimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbTipoRequerimiento.DisplayMember = "Descripcion";
			this.cmbTipoRequerimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRequerimiento.Location = new System.Drawing.Point(480, 21);
			this.cmbTipoRequerimiento.Name = "cmbTipoRequerimiento";
			this.cmbTipoRequerimiento.Size = new System.Drawing.Size(213, 21);
			this.cmbTipoRequerimiento.TabIndex = 1075;
			this.cmbTipoRequerimiento.ValueMember = "idRequerimientoTipo";
			this.cmbTipoRequerimiento.ValueChanged += new System.EventHandler(this.cmbTipoRequerimiento_ValueChanged_1);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(13, 28);
			this.label3.Name = "label3";
			this.label3.TabIndex = 1072;
			this.label3.Text = "* Asignado:";
			// 
			// cmbRequerimientoDepartamento
			// 
			this.cmbRequerimientoDepartamento.CausesValidation = false;
			this.cmbRequerimientoDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRequerimientoDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRequerimientoDepartamento.DataSource = this.udsRequerimientoDepartamento;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbRequerimientoDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbRequerimientoDepartamento.DisplayMember = "Descripcion";
			this.cmbRequerimientoDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRequerimientoDepartamento.Enabled = false;
			this.cmbRequerimientoDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRequerimientoDepartamento.Location = new System.Drawing.Point(113, 49);
			this.cmbRequerimientoDepartamento.MaxDropDownItems = 30;
			this.cmbRequerimientoDepartamento.Name = "cmbRequerimientoDepartamento";
			this.cmbRequerimientoDepartamento.Size = new System.Drawing.Size(249, 21);
			this.cmbRequerimientoDepartamento.TabIndex = 1071;
			this.cmbRequerimientoDepartamento.ValueMember = "idDepartamento";
			this.cmbRequerimientoDepartamento.Visible = false;
			// 
			// udsRequerimientoDepartamento
			// 
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn32.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn33.DefaultValue = "False";
			this.udsRequerimientoDepartamento.Band.Columns.AddRange(new object[] {
																																						 ultraDataColumn32,
																																						 ultraDataColumn33});
			// 
			// lblDepartamento
			// 
			this.lblDepartamento.AutoSize = true;
			this.lblDepartamento.BackColor = System.Drawing.Color.Transparent;
			this.lblDepartamento.Location = new System.Drawing.Point(13, 49);
			this.lblDepartamento.Name = "lblDepartamento";
			this.lblDepartamento.Size = new System.Drawing.Size(79, 16);
			this.lblDepartamento.TabIndex = 1070;
			this.lblDepartamento.Text = "Departamento:";
			this.lblDepartamento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblDepartamento.Visible = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(740, 14);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(34, 16);
			this.lblBodega.TabIndex = 1068;
			this.lblBodega.Text = "Local:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbResponsable
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance30;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Width = 150;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(113, 14);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(249, 21);
			this.cmbResponsable.TabIndex = 1037;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbResponsable_InitializeLayout);
			// 
			// cmbEmpresas
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance31;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresas.DataSource = this.udsEmpresa;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 304;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(460, 14);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(248, 21);
			this.cmbEmpresas.TabIndex = 1036;
			this.cmbEmpresas.ValueMember = "idempresa";
			this.cmbEmpresas.ValueChanged += new System.EventHandler(this.cmbEmpresas_ValueChanged);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(393, 14);
			this.label4.Name = "label4";
			this.label4.TabIndex = 1035;
			this.label4.Text = "* Empresa:";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(13, 14);
			this.label1.Name = "label1";
			this.label1.TabIndex = 1034;
			this.label1.Text = "* Responsable:";
			this.label1.Click += new System.EventHandler(this.label1_Click_1);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.udsProyecto;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Width = 312;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn38,
																											ultraGridColumn39});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(787, 14);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(248, 21);
			this.cmbBodega.TabIndex = 1069;
			this.cmbBodega.ValueMember = "Id";
			// 
			// udsProyecto
			// 
			ultraDataColumn34.DataType = typeof(int);
			this.udsProyecto.Band.Columns.AddRange(new object[] {
																														ultraDataColumn34,
																														ultraDataColumn35});
			// 
			// tabPage7
			// 
			this.tabPage7.Controls.Add(this.cmbSucursal);
			this.tabPage7.Controls.Add(this.label15);
			this.tabPage7.Controls.Add(this.cmbCiudad);
			this.tabPage7.Controls.Add(this.label14);
			this.tabPage7.Controls.Add(this.cmbCargoSolicitado);
			this.tabPage7.Controls.Add(this.label13);
			this.tabPage7.Location = new System.Drawing.Point(4, 22);
			this.tabPage7.Name = "tabPage7";
			this.tabPage7.Size = new System.Drawing.Size(1056, 112);
			this.tabPage7.TabIndex = 2;
			this.tabPage7.Text = "Cargo solicitado";
			this.tabPage7.Click += new System.EventHandler(this.tabPage7_Click);
			// 
			// cmbSucursal
			// 
			this.cmbSucursal.CausesValidation = false;
			this.cmbSucursal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataSource = this.udsRequerimientoBodegaXIdCiudad;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.Caption = "Descripción";
			ultraGridColumn41.Header.VisiblePosition = 2;
			ultraGridColumn41.Width = 244;
			ultraGridColumn42.Header.VisiblePosition = 3;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 4;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 6;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 7;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Hidden = true;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn40,
																											ultraGridColumn41,
																											ultraGridColumn42,
																											ultraGridColumn43,
																											ultraGridColumn44,
																											ultraGridColumn45,
																											ultraGridColumn46,
																											ultraGridColumn47});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbSucursal.DisplayMember = "name";
			this.cmbSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(104, 72);
			this.cmbSucursal.MaxDropDownItems = 30;
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.Size = new System.Drawing.Size(248, 21);
			this.cmbSucursal.TabIndex = 1073;
			this.cmbSucursal.ValueMember = "id";
			// 
			// udsRequerimientoBodegaXIdCiudad
			// 
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn36.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn37.DefaultValue = "False";
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			this.udsRequerimientoBodegaXIdCiudad.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn36,
																																								ultraDataColumn37,
																																								ultraDataColumn38,
																																								ultraDataColumn39,
																																								ultraDataColumn40,
																																								ultraDataColumn41,
																																								ultraDataColumn42});
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(8, 72);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(51, 16);
			this.label15.TabIndex = 1072;
			this.label15.Text = "Sucursal:";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.CausesValidation = false;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataSource = this.udsRequerimientoCiudadExp;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.Caption = "Descripción";
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn49.Width = 265;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Hidden = true;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn48,
																											ultraGridColumn49,
																											ultraGridColumn50});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbCiudad.DisplayMember = "name";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(104, 40);
			this.cmbCiudad.MaxDropDownItems = 30;
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(248, 21);
			this.cmbCiudad.TabIndex = 1071;
			this.cmbCiudad.ValueMember = "id";
			this.cmbCiudad.ValueChanged += new System.EventHandler(this.cmbCiudad_ValueChanged);
			// 
			// udsRequerimientoCiudadExp
			// 
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn43.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn44.DefaultValue = "False";
			this.udsRequerimientoCiudadExp.Band.Columns.AddRange(new object[] {
																																					ultraDataColumn43,
																																					ultraDataColumn44});
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(43, 16);
			this.label14.TabIndex = 1070;
			this.label14.Text = "Ciudad:";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCargoSolicitado
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCargoSolicitado.Appearance = appearance32;
			this.cmbCargoSolicitado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCargoSolicitado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargoSolicitado.DataSource = this.udsRequerimientoCargoCon;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 259;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn51,
																											ultraGridColumn52});
			this.cmbCargoSolicitado.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbCargoSolicitado.DisplayMember = "Descripcion";
			this.cmbCargoSolicitado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargoSolicitado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargoSolicitado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargoSolicitado.Location = new System.Drawing.Point(104, 8);
			this.cmbCargoSolicitado.Name = "cmbCargoSolicitado";
			this.cmbCargoSolicitado.Size = new System.Drawing.Size(248, 21);
			this.cmbCargoSolicitado.TabIndex = 1039;
			this.cmbCargoSolicitado.ValueMember = "idCargo";
			// 
			// udsRequerimientoCargoCon
			// 
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn45.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn46.DefaultValue = "False";
			this.udsRequerimientoCargoCon.Band.Columns.AddRange(new object[] {
																																				 ultraDataColumn45,
																																				 ultraDataColumn46});
			// 
			// label13
			// 
			this.label13.Location = new System.Drawing.Point(8, 8);
			this.label13.Name = "label13";
			this.label13.TabIndex = 1038;
			this.label13.Text = "Cargo solicitado:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(24, 0);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 16);
			this.label12.TabIndex = 821;
			this.label12.Text = "Numero";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// txtNumero
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance33;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(72, 0);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.ReadOnly = true;
			this.txtNumero.Size = new System.Drawing.Size(152, 21);
			this.txtNumero.TabIndex = 820;
			this.txtNumero.Visible = false;
			// 
			// frmSeguimiento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1186, 624);
			this.Controls.Add(this.tabDatosPrincipales);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.tabRequerimiento);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtNumero);
			this.Name = "frmSeguimiento";
			this.toolTip1.SetToolTip(this, "Guardar requerimiento");
			this.Load += new System.EventHandler(this.frmSeguimiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			this.tabRequerimiento.ResumeLayout(false);
			this.tabPage4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionCheckList)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoAreaAtencionActivosObtener)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionActivosObtener)).EndInit();
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotaIterna)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBusquedaMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsMenu)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.tabDatosPrincipales.ResumeLayout(false);
			this.tabPage5.ResumeLayout(false);
			this.grpAsignacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRequerimientoDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsProyecto)).EndInit();
			this.tabPage7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoBodegaXIdCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCiudadExp)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoSolicitado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCargoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
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
		string txtDirectorioarchivo = "";

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		// Declarar DataTable a nivel de la clase
		private DataTable dt = new DataTable();
		private DateTime _fechaAsignacion;

		private void frmSeguimiento_Load(object sender, System.EventArgs e)
		{
			// Equivalente a GETDATE()
			_fechaAsignacion = DateTime.Now;

				if(iTipo == iNuevo)
				{
					bNuevo = true;
					Title = "Nuevo requerimiento";
				}
				if(iTipo == iVer)
				{
					bVer = true;
					Title = "Ver requerimiento";
				}
	
				if(idRequerimientoDepartamento > 0)
				{
					sidRequerimientoDepartamento = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoDepartamentoModel '{0}',{1}","Descripcion",idRequerimientoDepartamento));
				}
				this.Text = Title;
				
				if(sidRequerimientoDepartamento != "")
				{
					this.Text += " (" + sidRequerimientoDepartamento+ ")";
				}

				//Carga id usuario tabla Inicio
				int idUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
				//Consultar el rol del usuario de ingreso
				int idrolUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}", idUsuarioInicio), false);
				// Cargar lista empresas vinculadas al usuario de ingreso
				DataSet dsEmpresas = Funcion.dsEjecucion(
					frmRequerimiento.sconexionInicio,
					String.Format(
					"EXEC RequerimientoRolEmpresa {0},{1}, {2}, {3}",
					idUsuarioInicio,
					idrolUsuarioInicio,
					idRequerimientoDepartamento,
					idEmpresa
					)
					);

				this.cmbEmpresas.DataSource = dsEmpresas;

				// --------------------------------------------------
				// ASIGNAR EMPRESA DE INGRESO (SI EXISTE)
				// --------------------------------------------------
				if (MenuLatinium.stIdDB > 0)
				{
					this.cmbEmpresas.Value = MenuLatinium.stIdDB;
				}

				// --------------------------------------------------
				// SI SOLO HAY UNA EMPRESA SELECCIONAR Y BLOQUEAR
				// --------------------------------------------------
				if (dsEmpresas != null &&
					dsEmpresas.Tables.Count > 0 &&
					dsEmpresas.Tables[0].Rows.Count == 1)
				{
					this.cmbEmpresas.Rows[0].Activate();
					this.cmbEmpresas.Enabled = false;
				}
				else
				{
					this.cmbEmpresas.Enabled = true;
				}
			//Cargar el nombre de la empresa que se logeo
				string sempresa = Funcion.sEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));
				//CargarResponsable
				this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
				this.cmbResponsable.Value = idResponsable;
				//Cargar el rol de la persona responsable
				RequerimientoRolUsuario();

			  CargarTipoRequerimiento();
				// Cargar Estado de Requerimiento
				CargarEstadoRequerimiento();
				CargarAsignado();

				//Cargar el grupo que pertenece la persona, pero por cada empresa
				//this.spnGrupo.Value = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.idGrupo From Usuarios e Where e.idUsuario = {0} Order By e.idUsuario", idUsuarioInicio),false);
				this.spnGrupo.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT TOP 1 idGrupo FROM USUARIO WHERE NOMBRE = '{0}'", MenuLatinium.stUsuario),false);
				//Cargar el menú seguridad, pero solo de los accesos que tiene cada grupo asignado.
				MenuSeguridad();
				ControlNuevo();
				RequerimientoRolPermisoCon();
				//cRAGAR 148.14
				//Ocultar pestaña panel si no es sistemas
				OcultarPestañaPanel();
				//Listado RequerimientoAreaAtencionActivosObtener
				RequerimientoAreaAtencionActivosObtener();
				//Cargar el departamento que pertenece
				RequerimientoGrupoCon();
				//iNICIALIZAR LA TABLA DOCUMENTO 
				InicializarTableDocumentos();
				ConfiguracionPorDepartamento();
		}
		private void CargarEstadoRequerimiento()
		{
			int getIdTipoRequerimiento = 0;

			if(this.cmbTipoRequerimiento.ActiveRow != null)
			{
				getIdTipoRequerimiento = (int)this.cmbTipoRequerimiento.Value;
			}

			DataSet dsEstado;

			// Validación básica
			if (idRequerimientoDepartamento <= 0)
			{
				this.cmbEstado.DataSource = null;
				this.cmbEstado.Enabled = false;
				return;
			}

			// Obtener estados desde SP pasando también el tipo de requerimiento
			dsEstado = Funcion.dsEjecucion(
				frmRequerimiento.sconexionPoint,
				String.Format(
				"Exec RequerimientoEstado {0}, {1}",
				idRequerimientoDepartamento,
				getIdTipoRequerimiento
				)
				);

			// Asignar datasource
			this.cmbEstado.DataSource = dsEstado;

			// Si solo hay un estado, seleccionarlo y bloquear
			if (dsEstado != null &&
				dsEstado.Tables.Count > 0 &&
				dsEstado.Tables[0].Rows.Count == 1)
			{
				this.cmbEstado.Rows[0].Activate();
				this.cmbEstado.Enabled = false;
			}
			else
			{
				this.cmbEstado.Enabled = true;
			}
		}


		private void CargarTipoRequerimiento()
		{
			// 1. Cargar tipos de requerimiento
			DataSet dsTipos = Funcion.dsEjecucion(
				frmRequerimiento.sconexionPoint,
				String.Format(
				"Exec RequerimientoTipoConsultar {0}, {1}",
				idRequerimientoDepartamento,
				idResponsable
				)
				);

			this.cmbTipoRequerimiento.DataSource = dsTipos;

			// Validación dura
			if (dsTipos == null ||
				dsTipos.Tables.Count == 0 ||
				dsTipos.Tables[0].Rows.Count == 0)
			{
				this.cmbTipoRequerimiento.SelectedRow = null;
				this.cmbTipoRequerimiento.Enabled = false;
				return;
			}

			// 2. Obtener tipo por default
			int idTipoDefault = Funcion.iEjecucion(
				frmRequerimiento.sconexionPoint,
				String.Format(
				"Exec RequerimientoTipoSeleccionDefault {0}, {1}",
				idRequerimientoDepartamento,
				idResponsable
				),
				false
				);

			// 3. Si no hay default salir
			if (idTipoDefault <= 0)
			{
				this.cmbTipoRequerimiento.SelectedRow = null;
				this.cmbTipoRequerimiento.Enabled = true;
				return;
			}

			// 4. Buscar por NOMBRE DE COLUMNA
			bool encontrado = false;
			int i;

			for (i = 0; i < this.cmbTipoRequerimiento.Rows.Count; i++)
			{
				if (this.cmbTipoRequerimiento.Rows[i].Cells.Exists("idRequerimientoTipo"))
				{
					if (Convert.ToInt32(
						this.cmbTipoRequerimiento.Rows[i]
						.Cells["idRequerimientoTipo"].Value
						) == idTipoDefault)
					{
						this.cmbTipoRequerimiento.Rows[i].Activate();
						this.cmbTipoRequerimiento.Value = idTipoDefault;
						encontrado = true;
						break;
					}
				}
			}

			// 5. Resultado final
			if (encontrado)
			{
				// Técnico bloquear
				this.cmbTipoRequerimiento.Enabled = false;
			}
			else
			{
				this.cmbTipoRequerimiento.SelectedRow = null;
				this.cmbTipoRequerimiento.Enabled = true;
			}
		}



		private void CargarAsignado()
		{
			DataSet dsAsignado = Funcion.dsEjecucion(
				frmRequerimiento.sconexionPoint,
				string.Format(
				"Exec RequerimientoAsignacionCon {0}, {1}, {2}, {3}",
				idResponsable,
				0,
				0,
				idRequerimientoDepartamento
				)
				);

			this.cmbAsignado.DataSource = dsAsignado;

			// --------------------------------------------
			// VALIDAR SI SOLO HAY UN REGISTRO
			// --------------------------------------------
			if (dsAsignado != null &&
				dsAsignado.Tables.Count > 0 &&
				dsAsignado.Tables[0].Rows.Count == 1)
			{
				this.cmbAsignado.Rows[0].Activate();
				this.cmbAsignado.Enabled = false;
			}
			else
			{
				this.cmbAsignado.Enabled = true;
			}

			//  APLICAR EXCEPCIÓN POR TURNO / SOPORTE
			AplicarAsignacionPorTurno();
		}

		private void AplicarAsignacionPorTurno()
		{
			int idAsignadoTurno;

			RequerimientoAsignacionTurnoSoporteService srv =
				new RequerimientoAsignacionTurnoSoporteService(
				new SqlConnection(frmRequerimiento.sconexionPoint)
				);

			// Debe existir tipo de requerimiento
			if (this.cmbTipoRequerimiento.Value == null)
				return;

			idAsignadoTurno = srv.ObtenerResponsableAsignado(
				idResponsable,
				Convert.ToInt32(this.cmbTipoRequerimiento.Value),
				idRequerimientoDepartamento
				);

			// Si NO hay asignación, no hacemos nada
			if (idAsignadoTurno <= 0)
			{
				return;
			}

			// ----------------------------------------
			// PASO 1: RECARGAR EL COMBO DE ASIGNADOS
			// ----------------------------------------
			CargarComboAsignado(); // <- ESTE ES EL PASO QUE FALTABA

			// Buscar SIEMPRE en el combo
			for (int i = 0; i < this.cmbAsignado.Rows.Count; i++)
			{
				if (Convert.ToInt32(this.cmbAsignado.Rows[i].Cells[0].Value) == idAsignadoTurno)
				{
					this.cmbAsignado.Rows[i].Activate();
					this.cmbAsignado.Value = idAsignadoTurno;
					this.cmbAsignado.Enabled = false;
					return;
				}
			}
		}

		private void CargarComboAsignado()
		{
			RequerimientoAsignacionTurnoSoporteService srv = new RequerimientoAsignacionTurnoSoporteService(
				new SqlConnection(frmRequerimiento.sconexionPoint)
				);

			this.cmbAsignado.DataSource = srv.ListarTecnicos(
				idRequerimientoDepartamento,
				Convert.ToInt32(this.cmbTipoRequerimiento.Value)
				);

			this.cmbAsignado.Refresh();
		}

		private void CargarCentroCosto()
		{
			try
			{
				System.Collections.ArrayList data =
					_requerimientoService.ObtenerCentrosCostoActivos();

				bool tieneDatos = (data != null && data.Count > 0);

				cmbBodega.Visible = tieneDatos;
				lblBodega.Visible = tieneDatos;

				if (!tieneDatos)
					return;

				cmbBodega.DataSource = data;
				cmbBodega.DisplayMember = "Nombre";
				cmbBodega.ValueMember = "Id";

				// --------------------------------------------------
				// 1 SI SOLO HAY UN REGISTRO SELECCIONAR Y BLOQUEAR
				// --------------------------------------------------
				if (data.Count == 1)
				{
					cmbBodega.Rows[0].Activate();
					cmbBodega.Enabled = false;
					return;
				}

				// --------------------------------------------------
				// 2 VALIDAR PROYECTO ASIGNADO AL USUARIO
				// --------------------------------------------------
				string squery = String.Format(
					"SELECT TOP 1 idUsuario FROM Usuario WHERE Nombre = '{0}'",
					MenuLatinium.stUsuario
					);

				int getidUsuario = Funcion.iEjecucion(
					getsconexionEmpresaSeguimiento,
					squery,
					false
					);

				string sQueryProyecto =
					"DECLARE @idProyecto INT " +
					"IF ( " +
					"   SELECT COUNT(1) " +
					"   FROM dbo.Bodega b " +
					"   INNER JOIN dbo.UsuariosBodegas ub ON ub.Bodega = b.Bodega " +
					"   WHERE ub.idUsuario = " + getidUsuario + " AND Factura = 1 " +
					") = 1 " +
					"BEGIN " +
					"   SELECT TOP 1 @idProyecto = p.idProyecto " +
					"   FROM Bodega b " +
					"   INNER JOIN UsuariosBodegas ub ON ub.Bodega = b.Bodega " +
					"   LEFT JOIN Proyecto p ON p.idProyecto = b.idProyecto " +
					"   WHERE ub.idUsuario = " + getidUsuario + " AND Remates = 0 " +
					"END " +
					"ELSE IF ( " +
					"   SELECT COUNT(1) " +
					"   FROM dbo.Bodega b " +
					"   INNER JOIN dbo.UsuariosBodegas ub ON ub.Bodega = b.Bodega " +
					"   WHERE ub.idUsuario = " + getidUsuario + " AND Factura = 1 " +
					") > 1 " +
					"BEGIN " +
					"   SELECT TOP 1 @idProyecto = ISNULL(n.idProyecto, 0) " +
					"   FROM Usuario u " +
					"   INNER JOIN Nomina n ON n.Codigo = u.Nombre " +
					"   WHERE u.idUsuario = " + getidUsuario + " " +
					"END " +
					"ELSE " +
					"BEGIN " +
					"   SET @idProyecto = 0 " +
					"END " +
					"SELECT @idProyecto";

				int idProyecto = Funcion.iEjecucion(
					getsconexionEmpresaSeguimiento,
					sQueryProyecto,
					false
					);

				// --------------------------------------------------
				// 3 SI TIENE PROYECTO ASIGNADO SELECCIONAR Y BLOQUEAR
				// --------------------------------------------------
				if (idProyecto > 0)
				{
					cmbBodega.Value = idProyecto;
					cmbBodega.Enabled = false;
				}
				else
				{
					cmbBodega.Enabled = true;
				}
			}
			catch (Exception)
			{
				cmbBodega.Visible = false;
				lblBodega.Visible = false;
			}
		}

		private void ConfiguracionPorDepartamento()
		{
			if(idRequerimientoDepartamento == frmRequerimiento.iRequerimientoDePersonal)
			{
				this.lblDepartamento.Text = "Área solicitante:";
				this.lblTitulo.Text = "Motivo vacante:";
				this.lblDescripcion.Text = "Observaciones adicionales:";
				this.cmbBodega.Visible = false;
				this.lblBodega.Visible = false;
				RequerimientoCargoCon();
				RequerimientoCiudadExp();
			}
		}

		private void RequerimientoCargoCon()
		{
			string squery = String.Format("Exec RequerimientoCargoCon");
			this.cmbCargoSolicitado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoCiudadExp()
		{
			string squery = String.Format("Exec RequerimientoCiudadExp");
			this.cmbCiudad.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoBodegaXIdCiudad()
		{
			this.cmbSucursal.Value = null;
			if (this.cmbCiudad.Value == null)
				return;

			string squery = "Exec RequerimientoBodegaXIdCiudad @idCiudad = " + ((int)this.cmbCiudad.Value).ToString();
			this.cmbSucursal.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void InicializarTableDocumentos()
		{
			// Inicializar estructura del DataTable (sin filas)
			dt.Columns.Add("txtDirectorioarchivo", typeof(string));

			// Asignar el DataSource al UltraGrid
			this.ugdDocumento.DataSource = dt;

			// Refrescar la grilla para asegurarse de que se actualice
			this.ugdDocumento.Refresh();

			// Asegurar que el grid permita agregar filas nuevas y mostrar la primera línea vacía
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;

			// Si deseas que la fila de adición aparezca arriba en lugar de abajo:
			// this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnTop;
		}
		

		private void RequerimientoGrupoCon()
		{
			int idTipoSeleccionado = 0; // Valor por defecto
			if (this.cmbTipoRequerimiento.Value != null)
			{
				idTipoSeleccionado = Convert.ToInt32(this.cmbTipoRequerimiento.Value);
			}

			if(RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado,"bDepartamento"))
			{
				this.cmbRequerimientoDepartamento.Visible = true;
				this.lblDepartamento.Visible = true;
				string getStUsuario = MenuLatinium.stUsuario;
				//Buscar el idDepartamento por el codigo 
				string squery = String.Format("Exec RequerimientoIdDepartamentoCon '{0}'", getStUsuario);
				this.cmbRequerimientoDepartamento.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
				// Verifica que haya filas en el control antes de seleccionar el primer registro
				if (this.cmbRequerimientoDepartamento.Rows.Count > 0)
				{
					this.cmbRequerimientoDepartamento.Rows[0].Selected = true;
				}
				else
				{
					this.cmbRequerimientoDepartamento.Visible = false;
					this.lblDepartamento.Visible = false;
				}
			}
		}

		// Definir un ArrayList global para almacenar los IDs seleccionados
		private System.Collections.ArrayList idsSeleccionados = new System.Collections.ArrayList();

		private void ugdRequerimientoAreaAtencionActivosObtener_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verificar si la columna que se actualizó es la columna de selección (ejemplo: "Seleccionar")
			if (e.Cell.Column.Key == "Seleccionar")
			{
				bool isSelected = (bool)e.Cell.Value;
				int idRequerimientoAreaAtencion = (int)e.Cell.Row.Cells["idRequerimientoAreaAtencion"].Value;
				int getYidRequerimientoDepartamento = idRequerimientoDepartamento;

				// Crear un ArrayList para almacenar los IDs seleccionados temporalmente
				ArrayList idsTempSeleccionados = new ArrayList();

				if (isSelected)
				{
					// Llamar al procedimiento almacenado para obtener idRequerimientoPerTecnico
					int idRequerimientoPerTecnico = ObtenerRequerimientoPerTecnicoPorID(idRequerimientoAreaAtencion, getYidRequerimientoDepartamento);

					// Si no hay selección previa, agregar el id actual
					if (idsSeleccionados.Count == 0)
					{
						idsSeleccionados.Add(idRequerimientoAreaAtencion);
						cmbAsignado.Value = idRequerimientoPerTecnico;  // Asignar el valor al combo
					}
					else
					{
						// Validar si el técnico actual es el mismo que el de la selección previa
						int idPerTecnicoActual = ObtenerRequerimientoPerTecnicoPorID((int)idsSeleccionados[0], getYidRequerimientoDepartamento);

						if (idRequerimientoPerTecnico != idPerTecnicoActual)
						{
							// Mostrar mensaje y revertir selección si los técnicos son diferentes
							MessageBox.Show("El requerimiento solo puede ser asignado a un Usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							e.Cell.Value = false;  // Deseleccionar la fila
							return;
						}
						else
						{
							// Si es el mismo técnico, agregar la fila seleccionada
							idsSeleccionados.Add(idRequerimientoAreaAtencion);
						}
					}
				}
				else
				{
					// Si la fila se deselecciona, eliminar el id de la lista
					idsSeleccionados.Remove(idRequerimientoAreaAtencion);
				}

				// Si no quedan seleccionados, resetear el combo
				if (idsSeleccionados.Count == 0)
				{
					cmbAsignado.Value = null;
				}
			}
		}

		private int ObtenerRequerimientoPerTecnicoPorID(int idRequerimientoAreaAtencion, int idRequerimientoDepartamento)
		{
			int idRequerimientoPerTecnico = 0;

			// Conexión a la base de datos y ejecución del procedimiento almacenado
			using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
			{
				conn.Open();
				using (SqlCommand cmd = new SqlCommand("RequerimientoAreaAtencionPerTecnicoPorID", conn))
				{
					cmd.CommandType = CommandType.StoredProcedure;
					cmd.Parameters.Add(new SqlParameter("@idRequerimientoAreaAtencion", idRequerimientoAreaAtencion));
					cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", idRequerimientoDepartamento));

					object result = cmd.ExecuteScalar();
					if (result != null)
					{
						idRequerimientoPerTecnico = Convert.ToInt32(result);
					}
				}
			}

			return idRequerimientoPerTecnico;
		}


		private void RequerimientoRolUsuario()
		{
			//Cual es el rol del usuario logeado
			idRequerimientoRol = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento), false);
			string squery = String.Format("Exec RequerimientoRolModel '{0}', {1}", "Subtitulo", idRequerimientoRol);
			sRequerimientoRol = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoRolPermisoCon()
		{
			try
			{
				//Permiso Valores técnicos
				bool bvalorestecnicos =
					Funcion.bEjecucion(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0},'{1}', {2}", 
					(int)this.cmbResponsable.Value,
					frmRequerimiento.iValoresTecnicos,
					idRequerimientoDepartamento
					), false); 

				if(bvalorestecnicos == false)
				{
					this.lblNotaInterna.Visible = false;
					this.txtNotaIterna.Visible = false;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				return;
			}
			
		}


		private void ControlNuevo()
		{
			if (bNuevo)
			{
				// Obtener el valor del tipo de requerimiento
				int idTipoRequerimiento = 0;
				int idCmbResponsable = 0;

				if (cmbTipoRequerimiento.ActiveRow != null && cmbTipoRequerimiento.Value != null)
				{
					idTipoRequerimiento = Convert.ToInt32(cmbTipoRequerimiento.Value);
				}
				if (cmbResponsable.ActiveRow != null && cmbResponsable.Value != null)
				{
					idCmbResponsable = (int)this.cmbResponsable.Value;
				}

				// Llamar al SP pasando idRequerimientoDepartamento y idTipoRequerimiento
				string squery = String.Format(
					"EXEC RequerimientoEstadoInicioXIdDepartamento {0}, {1}, {2}", 
					idRequerimientoDepartamento, 
					idTipoRequerimiento,
					idCmbResponsable
					);

				int idEstado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery, false);

				// -------------------------------
				// Seleccionar la fila correspondiente en cmbEstado
				// -------------------------------
				bool filaEncontrada = false;
				for (int i = 0; i < cmbEstado.Rows.Count; i++)
				{
					if (Convert.ToInt32(cmbEstado.Rows[i].Cells[0].Value) == idEstado)
					{
						cmbEstado.Rows[i].Activate();
						cmbEstado.Value = idEstado;
						filaEncontrada = true;
						break;
					}
				}

				// Si no encontró, asignar Value directamente (por si acaso)
				if (!filaEncontrada)
				{
					cmbEstado.Value = idEstado;
				}
			}
		}


		private void AplicarSeleccionMemoria()
		{
			if (_dtMenu == null)
				return;

			for (int i = 0; i < _dtMenu.Rows.Count; i++)
			{
				DataRow row = _dtMenu.Rows[i];

				if (row == null)
					continue;

				object val = row["id_seg"];
				if (val == null || val == DBNull.Value)
					continue;

				int idSeg;
				try
				{
					idSeg = Convert.ToInt32(val);
				}
				catch
				{
					continue;
				}

				// Marca o desmarca según el servicio
				if (_menuService != null && _menuService.Existe(idSeg))
				{
					row["Sel_ReqPantalla"] = true;
				}
				else
				{
					row["Sel_ReqPantalla"] = false;
				}
			}

			ugdMenu.DataBind(); // refresca el grid
		}

		private void MenuSeguridad()
		{
			MenuLatinium.MostrarModalProcesando();

			try
			{
				if (icontadoringreso == 0)
				{
					_dtMenu = FuncionesProcedimientos.dtGeneral(
						string.Format(
						"EXEC RequerimientoMenuSeguridadConsultar {0}, '{1}'",
						(int)this.spnGrupo.Value,
						this.txtBusquedaMenu.Text
						)
						);
				}
				else
				{
					_dtMenu = Funcion.dsEjecucion(
						getsconexionEmpresaSeguimiento,
						string.Format(
						"EXEC RequerimientoMenuSeguridadConsultar {0}, '{1}'",
						(int)this.spnGrupo.Value,
						this.txtBusquedaMenu.Text
						)
						).Tables[0]; // si dsEjecucion devuelve DataSet
				}

				if (_dtMenu == null || _dtMenu.Rows.Count == 0)
				{
					MenuLatinium.OcultarModalProcesando(); // importante si ya lo abriste
					return;
				}

				ugdMenu.DataSource = _dtMenu;
				ugdMenu.DataBind();

				AplicarSeleccionMemoria(); // primero aplica selección
				AplicarOrdenMenu();        // luego aplica orden
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
			finally
			{
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void AplicarOrdenMenu()
		{
			// 1 Validar grid y bandas
			if (this.ugdMenu == null)
				return;

			if (this.ugdMenu.DisplayLayout == null)
				return;

			if (this.ugdMenu.DisplayLayout.Bands.Count == 0)
				return;

			Infragistics.Win.UltraWinGrid.UltraGridBand band =
				this.ugdMenu.DisplayLayout.Bands[0];

			if (band == null)
				return;

			// 2 Validar que existan las columnas antes de ordenar
			if (!band.Columns.Exists("NombrePadre") ||
				!band.Columns.Exists("NombreHijo"))
			{
				// aquí simplemente no ordenas y evitas el crash
				return;
			}

			// 3 Aplicar orden SOLO si no existe ya
			if (!band.SortedColumns.Exists("NombrePadre"))
			{
				band.SortedColumns.Clear();

				band.SortedColumns.Add("NombrePadre", false, true);
				band.SortedColumns.Add("NombreHijo", false, true);
			}
		}

		private void ActivarCreacionRequerimiento()
		{
			bool bActivar = false;

			try
			{
				if (this.cmbAsignado.ActiveRow != null && this.cmbTipoRequerimiento.ActiveRow != null)
				{
					bActivar = true;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			if(bActivar)
			{
				this.tabRequerimiento.Enabled = true;
			}
			else
			{
				this.tabRequerimiento.Enabled = false;
			}

		}



		private void RequerimientoDocumentoCon(int Nuevo)
		{
			int setidAsignado = 0;
			try {setidAsignado = (int)this.cmbAsignado.Value;} 
			catch(Exception ex){string sms = ex.Message;}
				
			this.ugdDocumento.DataSource = 
				FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,string.Format("EXEC RequerimientoDocumentoCon {0}, {1}, {2}, '{3}'", 
				Nuevo, (int)this.cmbResponsable.Value, setidAsignado, Convert.ToDateTime(_fechaAsignacion).ToString("yyyyMMdd")
				)); 
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private bool _validacionOk;
		private void ValidacionGuardarRequerimiento()
		{
      _validacionOk = false;
			int idTipoSeleccionado = 0;

			// Verificar si el combo tiene valor seleccionado
			if (this.cmbTipoRequerimiento.Value != null)
			{
				idTipoSeleccionado = Convert.ToInt32(this.cmbTipoRequerimiento.Value);
			}

			#region Valida campos  
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione el responsable ")) return;
			if (!Validacion.vbComboVacio(this.cmbEmpresas, "Seleccione la empresa ")) return;
			// Llamar al SP enviando idRequerimientoDepartamento y tipo seleccionado
			if (RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Checklist"))
			{
				// Quitar la ordenación antes de actualizar los datos
				this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Bands[0].SortedColumns.Clear();

				#region Valida checklist 
				// Verificar si hay datos cargados en el grid
				if (this.ugdRequerimientoAreaAtencionActivosObtener.DataSource != null)
				{
					bool tieneSeleccion = false; // Variable para rastrear si hay alguna selección

					// Iterar sobre cada fila del grid
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.ugdRequerimientoAreaAtencionActivosObtener.Rows)
					{
						// Obtener el ID de la fila actual
						int idRequerimientoAreaAtencion = (int)row.Cells["idRequerimientoAreaAtencion"].Value;
						// Verificar si el ID está en la lista de seleccionados
						if (idsSeleccionados.Contains(idRequerimientoAreaAtencion))
						{
							tieneSeleccion = true; // Al menos un valor seleccionado
						}
					}

					// Si no hay ningún elemento seleccionado, mostrar mensaje
					if (!tieneSeleccion)
					{
						// Verificar si hay filas y aplicar la ordenación nuevamente
						if (this.ugdRequerimientoAreaAtencionActivosObtener.Rows.Count > 0)
						{
							// Volver a aplicar la ordenación a la columna "AreaAtencionTipo"
							this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Bands[0].SortedColumns.Add("AreaAtencionTipo", false, true);
						}
						MessageBox.Show("Debe seleccionar al menos un ítem en el checklist.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
				}
				#endregion Valida checklist

				//Validar descripción del checklist
				if (!Validacion.vbTexto(this.txtDescripcionCheckList, 20, 2000, "Ingrese una descripción")) return;
				this.txtDescripcion.Text = this.txtDescripcionCheckList.Text;
			}

			if (RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "AsignacionPorDisponibilidad"))
			{
				this.cmbAsignado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, String.Format("Exec RequerimientoPerTecnicoCon @idRequerimientoDepartamento = {0}", idRequerimientoDepartamento));
				string AsignadoObtener = String.Format("Exec RequerimientoAsignacionPorDisponibilidad {0}", idRequerimientoDepartamento);
				this.cmbAsignado.Value = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, AsignadoObtener, false);
			}
			else
			{
				if (!Validacion.vbComboVacio(this.cmbAsignado, "Seleccione el asignado")) return;
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega")) return;
			}

			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) return;
			if (!Validacion.vbComboVacio(this.cmbTipoRequerimiento, "Seleccione el tipo de requerimiento ")) return;


			if(RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado ,"bDescripcion"))
			{
				if(frmRequerimiento.iRequerimientoDePersonal == idRequerimientoDepartamento)
				{
					if (!Validacion.vbTexto(this.txtTitulo, 10, 50, "Ingrese motivo de la vacante")) return;
					if (!Validacion.vbTexto(this.txtDescripcion, 20, 2000, "Ingrese observaciones adicionales")) return;
				}
				else
				{
					if (!Validacion.vbTexto(this.txtTitulo, 10, 50, "Ingrese un título")) return;
					if (!Validacion.vbTexto(this.txtDescripcion, 20, 2000, "Ingrese una descripción")) return;
				}

			}

			if(RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "bDepartamento") && this.cmbRequerimientoDepartamento.Rows.Count > 0)
			{
				if (!Validacion.vbComboVacio(this.cmbRequerimientoDepartamento, "Ingrese el departamento")) return;
			}

			if (!Validacion.vbCampoEnteroVacio(this.spnGrupo,"Ingrese el grupo", 1, 100)) return;

			if(RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "CargoSolicitado"))
			{
				if (!Validacion.vbComboVacio(this.cmbCargoSolicitado, "Ingrese el cargo solicitad")) return;
				if (!Validacion.vbComboVacio(this.cmbCiudad, "Ingrese la ciudad")) return;
				if (!Validacion.vbComboVacio(this.cmbSucursal, "Ingrese la sucursal")) return;
			}

			if (RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Pantalla"))
			{
				#region Valida menú seguridad

				if (_dtMenu == null || _dtMenu.Rows.Count == 0)
				{
					MessageBox.Show("Proceso Finalizado, no hay datos en el menú.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				int countSeleccionados = 0;

				for (int i = 0; i < _dtMenu.Rows.Count; i++)
				{
					DataRow row = _dtMenu.Rows[i];
					if (row == null)
						continue;

					object val = row["Sel_ReqPantalla"];
					if (val != null && val != DBNull.Value)
					{
						bool seleccionado = false;
						try
						{
							seleccionado = Convert.ToBoolean(val);
						}
						catch
						{
							continue; // ignora si no se puede convertir
						}

						if (seleccionado)
							countSeleccionados++;
					}
				}

				if (countSeleccionados == 0)
				{
					MessageBox.Show("Proceso Finalizado, se debe seleccionar una pantalla",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				if (countSeleccionados > 1)
				{
					MessageBox.Show("Proceso Finalizado, se debe seleccionar una sola pantalla",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				#endregion
			}


			#endregion Valida campos  

			#region Valida documentos
			if(bRequerimientosDocumento)
			{
				// Validación si se ha ingresado al menos un documentos eso solo sería de hacer un select en la tabla documentosrequerimiento con el estado 1
				int ugdDocumentotamano = this.ugdDocumento.Rows.Count;
				if(!RequerimientoClass.ValidacionDocumentos((int)this.cmbEstado.Value, idRequerimientoDepartamento, ugdDocumentotamano))return;
			}
			#endregion Valida documentos 

			#region Mensaje de confirmacion 

			string Mensaje = String.Format("¿Esta seguro de grabar el requerimiento para el {0}?.", sidRequerimientoDepartamento);
			if	(DialogResult.Yes == MessageBox.Show(Mensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				this.DialogResult = DialogResult.OK;
				this.Close();
			}
			else return;
			#endregion Mensaje de confirmacion 

			// si llega hasta el final:
			_validacionOk = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			ValidacionGuardarRequerimiento();
			if (!_validacionOk)
				return; // SE DETIENE TODO AQUÍ

			int idTipoSeleccionado = 0;

			// Verificar si el combo tiene valor seleccionado
			if (this.cmbTipoRequerimiento.Value != null)
			{
				idTipoSeleccionado = Convert.ToInt32(this.cmbTipoRequerimiento.Value);
			}
			
			#region	Generar rutas del documento
			string squery = "";
			string squery2 = "";
			string squery3 = "";

			#region Selección Pantalla
			int idPantalla = 0;
			string modulo = "";
			string pantalla = "";
			string subPantalla = "";

			for (int i = 0; i < _dtMenu.Rows.Count; i++)
			{
				DataRow row = _dtMenu.Rows[i];
				if (row == null) continue;

				bool seleccionado = false;
				try { seleccionado = Convert.ToBoolean(row["Sel_ReqPantalla"]); }
				catch { continue; }

				if (seleccionado)
				{
					try { idPantalla = Convert.ToInt32(row["id_seg"]); } 
					catch {}
					try { modulo = Convert.ToString(row["NombrePadre"]); } 
					catch {}
					try { pantalla = Convert.ToString(row["NombreHijo"]); } 
					catch {}
					try { subPantalla = Convert.ToString(row["NombreSubHijo"]); } 
					catch {}
					break; // SOLO UNA
				}
			}
			#endregion Selección Pantalla
				
			using (SqlConnection oConexion = new SqlConnection(frmRequerimiento.sconexionPoint)) //("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
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

					#region Numeracion
					// ===========================
					//   NUMERACIÓN
					// ===========================
					if (bNuevo)
					{
						try
						{
							int idTipoFactura = 0;
							object result = null;
							object num = null;

							// LIMPIAR parámetros previos
							oCmdActualiza.Parameters.Clear();

							// 1) Obtener idTipoFactura
							oCmdActualiza.CommandText = "RequerimientoDepartamentoModel";
							oCmdActualiza.CommandType = CommandType.StoredProcedure;

							oCmdActualiza.Parameters.Add("@NombreColumna", SqlDbType.VarChar, 50).Value = "idTipoFactura";
							oCmdActualiza.Parameters.Add("@id", SqlDbType.Int).Value = idRequerimientoDepartamento;

							result = oCmdActualiza.ExecuteScalar();

							if (result != null && result != DBNull.Value)
								idTipoFactura = Convert.ToInt32(result);
							else
							{
								oTransaction.Rollback();
								MessageBox.Show("No fue posible obtener el tipo de factura.", 
									"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								return;
							}

							// =======================
							//     NUMERACION
							// =======================

							oCmdActualiza.Parameters.Clear(); // limpiar antes del SP
							oCmdActualiza.CommandText = "RequermientoNumeracion";
							oCmdActualiza.CommandType = CommandType.StoredProcedure;

							oCmdActualiza.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = idTipoFactura;

							num = oCmdActualiza.ExecuteScalar();

							if (num == null || num == DBNull.Value)
							{
								oTransaction.Rollback();

								MessageBox.Show(
									"No fue posible generar numeración. Intente nuevamente.",
									"Point Technology",
									MessageBoxButtons.OK,
									MessageBoxIcon.Error
									);

								Funcion.LogSistema(
									frmRequerimiento.sconexionInicio,
									this.Name + ": " + this.Text,
									"ExecuteScalar devolvió NULL al generar numeración",
									"",
									Funcion.slogtipoerror
									);

								return;
							}

							this.txtNumero.Text = Convert.ToString(num);
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							Funcion.LogSistema(
								frmRequerimiento.sconexionInicio,
								this.Name + ": " + this.Text,
								"Commit Exception Type : " + ex.GetType().ToString() + " " + ex.Message,
								"",
								Funcion.slogtipoerror
								);

							MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
					}
					#endregion Numeracion

					#region Grabar requerimiento

					int idSeguimiento = 0;
					int getXBodega = 0;
					string getXsBodega = "";
					int getXDepartamento = 0;
					string getXsDepartamento = "";
					int getidRequerimiento = 0;
					int getXIdCargo = 0;
					string getXsIdCargo = "";
					int getXIdCiudad = 0;
					string getXsIdCiudad = "";

					// Obtener valores sin romper la ejecución
					try { getXDepartamento = (int)this.cmbRequerimientoDepartamento.Value; } 
					catch {}
					try { getXsDepartamento = this.cmbRequerimientoDepartamento.Text; } 
					catch {}

					if (idRequerimientoDepartamento == frmRequerimiento.iRequerimientoDePersonal)
					{
						try { getXBodega = (int)this.cmbSucursal.Value; } 
						catch {}
						try { getXsBodega = this.cmbSucursal.Text; } 
						catch {}
					}
					else
					{
						try { getXBodega = (int)this.cmbBodega.Value; } 
						catch {}
						try { getXsBodega = this.cmbBodega.Text; } 
						catch {}
					}

					try { getXIdCargo = (int)this.cmbCargoSolicitado.Value; } 
					catch {}
					try { getXsIdCargo = this.cmbCargoSolicitado.Text; } 
					catch {}
					try { getXIdCiudad = (int)this.cmbCiudad.Value; } 
					catch {}
					try { getXsIdCiudad = this.cmbCiudad.Text; } 
					catch {}

					try
					{
						string setDescripcion = RequerimientoClass.ControlErrorComillasSimples(this.txtDescripcion.Value);
						string setTitulo = RequerimientoClass.ControlErrorComillasSimples(this.txtTitulo.Value);
						string setNotaInterna = RequerimientoClass.ControlErrorComillasSimples(this.txtNotaIterna.Value);

						// =====================================================
						// 1) GUARDAR REQUERIMIENTO - .NET 1.1
						// =====================================================

						SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint);
						SqlCommand cmd = new SqlCommand("RequerimientoCrear", conn);
						cmd.CommandType = CommandType.StoredProcedure;

						cmd.Parameters.Add("@idEmpresa", SqlDbType.Int).Value = (int)this.cmbEmpresas.Value;
						cmd.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = setDescripcion;
						cmd.Parameters.Add("@idEstado", SqlDbType.Int).Value = (int)this.cmbEstado.Value;
						cmd.Parameters.Add(
							"@FechaTentativaEntrega",
							SqlDbType.DateTime
							).Value = DateTime.Now;
						cmd.Parameters.Add("@NotaInterna", SqlDbType.VarChar).Value = setNotaInterna;
						cmd.Parameters.Add("@idPersonal", SqlDbType.Int).Value = (int)this.cmbResponsable.Value;
						cmd.Parameters.Add("@Titulo", SqlDbType.VarChar, 50).Value = setTitulo;
						cmd.Parameters.Add("@idRequerimientoTipo", SqlDbType.Int).Value = (int)this.cmbTipoRequerimiento.Value;
						cmd.Parameters.Add("@idAsignado", SqlDbType.Int).Value = (int)this.cmbAsignado.Value;
						cmd.Parameters.Add("@Numero", SqlDbType.VarChar, 50).Value = this.txtNumero.Text;
						cmd.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = _fechaAsignacion;
						cmd.Parameters.Add("@Responsable", SqlDbType.VarChar, 500).Value = this.cmbResponsable.Text;
						cmd.Parameters.Add("@Asignado", SqlDbType.VarChar, 500).Value = this.cmbAsignado.Text;
						cmd.Parameters.Add("@Empresa", SqlDbType.VarChar, 500).Value = this.cmbEmpresas.Text;
						cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;

						// Extras
						cmd.Parameters.Add("@Bodega", SqlDbType.Int).Value = getXBodega;
						cmd.Parameters.Add("@NombreBodega", SqlDbType.VarChar, 500).Value = getXsBodega;
						cmd.Parameters.Add("@idDepartamento", SqlDbType.Int).Value = getXDepartamento;
						cmd.Parameters.Add("@NombreDepartamento", SqlDbType.VarChar, 800).Value = getXsDepartamento;
						cmd.Parameters.Add("@idCargo", SqlDbType.Int).Value = getXIdCargo;
						cmd.Parameters.Add("@NombreCargo", SqlDbType.VarChar, 500).Value = getXsIdCargo;
						cmd.Parameters.Add("@idCiudad", SqlDbType.Int).Value = getXIdCiudad;
						cmd.Parameters.Add("@NombreCiudad", SqlDbType.VarChar, 500).Value = getXsIdCiudad;

						// ================================
						// CAMPOS NUEVOS DESDE FRONT
						// ================================
						cmd.Parameters.Add("@idPantalla", SqlDbType.Int).Value =
							(idPantalla > 0) ? (object)idPantalla : DBNull.Value;

						SqlParameter pModulo = cmd.Parameters.Add("@Modulo", SqlDbType.VarChar, 500);
						if (modulo == null || modulo.Trim() == "")
							pModulo.Value = DBNull.Value;
						else
							pModulo.Value = modulo;

						SqlParameter pPantalla = cmd.Parameters.Add("@Pantalla", SqlDbType.VarChar, 500);
						if (pantalla == null || pantalla.Trim() == "")
							pPantalla.Value = DBNull.Value;
						else
							pPantalla.Value = pantalla;

						SqlParameter pSubPantalla = cmd.Parameters.Add("@SubPantalla", SqlDbType.VarChar, 500);
						if (subPantalla == null || subPantalla.Trim() == "")
							pSubPantalla.Value = DBNull.Value;
						else
							pSubPantalla.Value = subPantalla;

						// Salida
						SqlParameter pOut = new SqlParameter("@idSeguimiento", SqlDbType.Int);
						pOut.Direction = ParameterDirection.Output;
						cmd.Parameters.Add(pOut);

						conn.Open();
						cmd.ExecuteNonQuery();

						idSeguimiento = (pOut.Value != null) ? (int)pOut.Value : 0;

						conn.Close();
						conn.Dispose();
						cmd.Dispose();

						// =====================================================
						// 2) Recuperar idRequerimiento
						// =====================================================

						oCmdActualiza.Parameters.Clear();
						oCmdActualiza.CommandText = "SeguimientoModel";
						oCmdActualiza.CommandType = CommandType.StoredProcedure;

						oCmdActualiza.Parameters.Add("@NombreColumna", SqlDbType.VarChar, 50).Value = "idRequerimiento";
						oCmdActualiza.Parameters.Add("@idSeguimiento", SqlDbType.Int).Value = idSeguimiento;

						object retorno = oCmdActualiza.ExecuteScalar();
						if (retorno != null && retorno != DBNull.Value)
							getidRequerimiento = Convert.ToInt32(retorno);

						// =====================================================
						// 3) CHECKLIST
						// =====================================================

						if (RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Checklist"))
						{
							int i = 0;
							for (i = 0; i < this.ugdRequerimientoAreaAtencionActivosObtener.Rows.Count; i++)
							{
								Infragistics.Win.UltraWinGrid.UltraGridRow row =
									this.ugdRequerimientoAreaAtencionActivosObtener.Rows[i];

								bool seleccionado = false;
								try { seleccionado = (bool)row.Cells["Seleccionar"].Value; } 
								catch {}

								if (seleccionado)
								{
									int idArea = (int)row.Cells["idRequerimientoAreaAtencion"].Value;

									oCmdActualiza.Parameters.Clear();
									oCmdActualiza.CommandText = "RequerimientoAreaAtencionAsignaInsertar";
									oCmdActualiza.CommandType = CommandType.StoredProcedure;

									oCmdActualiza.Parameters.Add("@idRequerimiento", SqlDbType.Int).Value = getidRequerimiento;
									oCmdActualiza.Parameters.Add("@Estado", SqlDbType.Bit).Value = 1;
									oCmdActualiza.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;
									oCmdActualiza.Parameters.Add("@idRequerimientoAreaAtencion", SqlDbType.Int).Value = idArea;

									oCmdActualiza.ExecuteNonQuery();
								}
							}
						}
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						Funcion.LogSistema(
							frmRequerimiento.sconexionInicio,
							this.Name + ": " + this.Text,
							"Commit Exception Type : " + ex.GetType().ToString() + " " + ex.Message,
							"",
							Funcion.slogtipoerror
							);

						MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					#endregion

					// Supongamos que estos datos ya los tienes en tu form
					string usuario = MenuLatinium.stUsuario;
					string cadenaConexion = frmRequerimiento.sconexionPoint;
					string dirFacturacion = MenuLatinium.stDirFacturacion;
					// Instancia del servicio
					DocumentoService docService = new DocumentoService(usuario, cadenaConexion, dirFacturacion);
					// Crear lista de documentos desde el grid
					ArrayList listaDocumentos = new ArrayList();
					for (int i = 0; i < this.ugdDocumento.Rows.Count; i++)
					{
						UltraGridRow row = this.ugdDocumento.Rows[i];

						DocumentoModel doc = new DocumentoModel();
						doc.txtDirectorioarchivo = row.Cells["txtDirectorioarchivo"].Value.ToString();
						doc.idResponsable = Convert.ToInt32(this.cmbResponsable.Value);
						doc.idAsignado = Convert.ToInt32(this.cmbAsignado.Value);

						listaDocumentos.Add(doc);
					}

						bool ok = docService.GuardarDocumentos(
							oCmdActualiza,                 // tu SqlCommand ya conectado y con transacción
							idSeguimiento,                 // id del seguimiento que obtuviste al crear el requerimiento
							txtNumero.Text,                // número de requerimiento
							sidRequerimientoDepartamento,  // departamento
							listaDocumentos,               // la lista de documentos que creaste
							_fechaAsignacion               // fecha de asignación
							);

						if (!ok)
						{
							oTransaction.Rollback();
							MessageBox.Show("No fue posible guardar los documentos. Verifique e intente nuevamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}

					// FIN DE FOR
					oTransaction.Commit();
					oConexion.Close();

					#region Enviar notificacion
					frmRequerimiento.RequerimientoNotificacionCrear(
						(int)this.cmbResponsable.Value, (int)this.cmbAsignado.Value, frmRequerimiento.iRequerimientoMensajeNuevo, idSeguimiento, 1, idRequerimientoDepartamento
						);
					#endregion Enviar notificacion

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se generó el siguiente número de requerimiento: {0}", this.txtNumero.Text), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;

				}
				catch (Exception ex)
				{
					this.txtNumero.Text = "";
					oTransaction.Rollback();
					MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					//Funcion.WriteLog(ex.Message);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery + " || "+ squery2 + " || " + squery3, Funcion.slogtipoerror);
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


		private RequerimientoMenuMemoriaService _menuService =
			new RequerimientoMenuMemoriaService();
		private DataTable _dtMenu; // variable de clase para almacenar la tabla

		private void ugdMenu_AfterCellUpdate(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key != "Sel_ReqPantalla") return;

			bool valorActual = Convert.ToBoolean(e.Cell.Value);
			bool valorAnterior = Convert.ToBoolean(e.Cell.OriginalValue);

			if (valorActual == valorAnterior) return;

			ProcesarSeleccionMenu(e);
		}

		private void ProcesarSeleccionMenu(
			Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			_menuService.ProcesarSeleccion(
				(int)e.Cell.Row.Cells["id_seg"].Value,
				(bool)e.Cell.Row.Cells["Sel_ReqPantalla"].Value,
				e.Cell.Row.Cells["NombrePadre"].Value.ToString(),
				e.Cell.Row.Cells["NombreHijo"].Value.ToString(),
				e.Cell.Row.Cells["NombreSubHijo"].Value == DBNull.Value
				? ""
				: e.Cell.Row.Cells["NombreSubHijo"].Value.ToString()
				);
		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

//		private void ugdDocumento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
//		{
//			long itamano = 0;
//			bool bResultado = false;
//			string sMensaje = "";
//			DataSet dsConfiguracion = new DataSet();
//			string txtDirectorioarchivo = "";
//			if (e.Cell.Column.Key == "btnDocumento") // Agregar documento
//			{
//				try
//				{
//					OpenFileDialog buscar = new OpenFileDialog();
//					if (buscar.ShowDialog() == DialogResult.OK)
//					{
//						txtDirectorioarchivo = buscar.FileName;
//						System.IO.FileInfo info = new System.IO.FileInfo(txtDirectorioarchivo);
//						itamano = info.Length;
//						dsConfiguracion = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, 
//							String.Format("EXEC RequerimientoConfgVal {0}, {1}, {2}", 
//							frmRequerimiento.iConfiguracionDocumento, itamano, frmRequerimiento.iConfiguracionTamano));
//						bResultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
//						sMensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();
//						if (bResultado)
//						{
//							// Obtener la fuente de datos actual de la grilla
//							DataTable dt = (DataTable)this.ugdDocumento.DataSource;
//							if (dt != null)
//							{
//								// Crear una nueva fila y asignarle el valor del archivo
//								DataRow nuevaFila = dt.NewRow();
//								nuevaFila["txtDirectorioarchivo"] = txtDirectorioarchivo;
//								// Insertar la nueva fila en la primera posición
//								dt.Rows.InsertAt(nuevaFila, 0);
//								// Actualizar la grilla
//								this.ugdDocumento.DataSource = dt;
//								this.ugdDocumento.Refresh();
//								// Asegurarse de que la primera fila sea visible
//								this.ugdDocumento.Rows[0].Activate();
//							}
//						}
//						else
//						{
//							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							return;
//						}
//					}
//				}
//				catch (Exception ex)
//				{
//					MessageBox.Show("Error al seleccionar archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				}
//			}
//			else if (e.Cell.Column.Key == "btnVer") // Ver documento
//			{
//				try
//				{
//					txtDirectorioarchivo = e.Cell.Row.Cells["txtDirectorioarchivo"].Value.ToString();
//					if (!(txtDirectorioarchivo == null || txtDirectorioarchivo == ""))
//					{
//						Process.Start(txtDirectorioarchivo);
//					}
//					else
//					{
//						MessageBox.Show("El archivo no existe en la ruta especificada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//					}
//				}
//				catch (Exception ex)
//				{
//					MessageBox.Show("Error al abrir archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				}
//			}
//			else if (e.Cell.Column.Key == "btnEliminar") // Eliminar documento de la grilla
//			{
//				try
//				{
//					if (MessageBox.Show("¿Está seguro de eliminar este documento de la lista?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
//					{
//						e.Cell.Row.Delete(false);
//					}
//				}
//				catch (Exception ex)
//				{
//					MessageBox.Show("Error al eliminar archivo de la lista: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				}
//			}
//		}
//
//
		private bool ValidarDocumentoSeleccionado(
			string rutaArchivo,
			out string mensaje)
		{
			return _documentoService.AgregarDocumento(
				rutaArchivo,
				frmRequerimiento.sconexionPoint,
				frmRequerimiento.iConfiguracionDocumento,
				frmRequerimiento.iConfiguracionTamano,
				out mensaje
				);
		}

		private bool ConfirmarEliminacionDocumento()
		{
			return MessageBox.Show(
				"¿Está seguro de eliminar este documento de la lista?",
				"Confirmación",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Question
				) == DialogResult.Yes;
		}


		private void ugdDocumento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Asegurarse de que el DataSource exista
			DataTable dt = (DataTable)this.ugdDocumento.DataSource;
			if (dt == null)
			{
				MessageBox.Show("El UltraDataSource no está inicializado.", "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Agregar la columna oculta DocumentoRef si no existe
			if (!dt.Columns.Contains("DocumentoRef"))
			{
				dt.Columns.Add("DocumentoRef", typeof(object));

				// Ocultar columna en el grid
				this.ugdDocumento.DisplayLayout.Bands[0].Columns["DocumentoRef"].Hidden = true;
			}

			// ===================== AGREGAR DOCUMENTO =====================
			if (e.Cell.Column.Key == "btnDocumento")
			{
				string mensaje = "";
				using (OpenFileDialog buscar = new OpenFileDialog())
				{
					if (buscar.ShowDialog() == DialogResult.OK)
					{
						// Validar documento
						bool resultado = ValidarDocumentoSeleccionado(buscar.FileName, out mensaje);
						if (!resultado)
						{
							MessageBox.Show(mensaje, "Point Technology",
								MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}

						// Crear modelo y guardar en el service
						DocumentoModel doc = new DocumentoModel();
						doc.txtDirectorioarchivo = buscar.FileName;
						_documentoService.ObtenerDocumentos().Add(doc);

						// Agregar fila al DataTable/UltraDataSource
						DataRow fila = dt.NewRow();
						fila["txtDirectorioarchivo"] = doc.txtDirectorioarchivo;
						fila["DocumentoRef"] = doc;  // columna oculta con objeto completo
						dt.Rows.InsertAt(fila, 0);

						this.ugdDocumento.Refresh();
						this.ugdDocumento.Rows[0].Activate();
					}
				}
			}
				// ===================== VER DOCUMENTO =====================
			else if (e.Cell.Column.Key == "btnVer")
			{
				try
				{
					DocumentoModel doc = (DocumentoModel)e.Cell.Row.Cells["DocumentoRef"].Value;
					string mensaje;
					if (!_documentoService.VerDocumento(doc, out mensaje))
					{
						MessageBox.Show(mensaje, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al abrir documento: " + ex.Message, "Error",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
				// ===================== ELIMINAR DOCUMENTO =====================
			else if (e.Cell.Column.Key == "btnEliminar")
			{
				try
				{
					if (!ConfirmarEliminacionDocumento())
						return;

					DocumentoModel doc = (DocumentoModel)e.Cell.Row.Cells["DocumentoRef"].Value;
					string mensaje;
					if (!_documentoService.EliminarDocumento(doc, out mensaje))
					{
						MessageBox.Show(mensaje, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}

					// Eliminar visualmente del grid
					e.Cell.Row.Delete(false);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al eliminar archivo de la lista: " + ex.Message, "Error",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void ugdMenu_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
			//MenuSeguridad();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			MenuSeguridad();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void label9_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ugdMenu_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbAsignado_ValueChanged(object sender, System.EventArgs e)
		{
			//RequerimientoDocumentoCon(iNuevo);
			//ActivarCreacionRequerimiento();
		}

		private void cmbAsignado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			//ActivarCreacionRequerimiento();
		}

		private void cmbTipoRequerimiento_ValueChanged(object sender, System.EventArgs e)
		{
			//ActivarCreacionRequerimiento();
		}

		private void cmbTipoRequerimiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			//ActivarCreacionRequerimiento();
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private RequerimientoService _requerimientoService;
		private string _conexionEmpresaActual;

		private void cmbEmpresas_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEmpresas.ActiveRow != null)
			{
				icontadoringreso = 1;

				int idEmpresa = (int)this.cmbEmpresas.Value;

				// Obtener la conexión de la empresa seleccionada
				_conexionEmpresaActual = Funcion.ObtenerCadenaConexionEmpresa(
					idEmpresa,
					frmRequerimiento.sconexionInicio,
					sconexionEmpresaSeguimiento
					);
				getsconexionEmpresaSeguimiento = _conexionEmpresaActual;

				// RECREAR EL SERVICE CON LA NUEVA CONEXIÓN
				_requerimientoService = new RequerimientoService(_conexionEmpresaActual);

				MenuSeguridad();
				CargarCentroCosto();
				RequerimientoCargoCon();
			}
		}

		private void txtBusquedaMenu_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.ValidaLetraDigito(sender,e, ' ');
		}

		private void OcultarPestañaPanel()
		{
			// Obtener el tipo seleccionado del combo
			int idTipoSeleccionado = 0;
			if (this.cmbTipoRequerimiento.Value != null)
				idTipoSeleccionado = Convert.ToInt32(this.cmbTipoRequerimiento.Value);


			// Verifica si el idRequerimientoDepartamento es diferente del valor esperado
			if (!RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Pantalla"))
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
			if (!RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Checklist"))
			{
				// Oculta tabPage4
				if (this.tabRequerimiento.TabPages.Contains(tabPage4))
				{
					tabRequerimiento.TabPages.Remove(tabPage4); // Remueve tabPage4 del TabControl
				}
			}
			else
			{
				// Asegúrate de que tabPage2 esté visible si la condición no se cumple
				if (!tabRequerimiento.TabPages.Contains(tabPage4))
				{
					tabRequerimiento.TabPages.Add(tabPage4); // Añade tabPage4 de nuevo al TabControl
				}
				//Mostrar el nuevo item solo si se le activa en requerimientopertecnico el usuario
				bool NuevoItemCheckList = false;
				//COnsolutar el id
				string squery1 = string.Format("EXEC RequerimientoPerTecnicoXIdPersonal {0}, {1}", 
					idRequerimientoDepartamento, 
					(int)this.cmbResponsable.Value
					);
				int getIdRequerimientoPerTecnico = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery1, false);
				string squery2 = string.Format("EXEC RequerimientoPerTecnicoModel '{0}', {1}", "NuevoItemCheckList", getIdRequerimientoPerTecnico);
				NuevoItemCheckList = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, squery2, false);
					if(NuevoItemCheckList)
					{
						this.btnNuevo.Visible = true;
					}
			}

			// Verifica si el idRequerimientoDepartamento es diferente del valor esperado
			if (!RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "bDescripcion"))
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

			if (RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "Asignacion"))
			{
				this.grpAsignacion.Visible = false;
			}
			else
			{
					this.grpAsignacion.Visible = true;
			}

			if (!RequerimientoDepartamentoTabMostrar(idRequerimientoDepartamento, idTipoSeleccionado, "CargoSolicitado"))
			{
				// Oculta tabPage4
				if (this.tabDatosPrincipales.TabPages.Contains(tabPage7))
				{
					tabDatosPrincipales.TabPages.Remove(tabPage7); // Remueve tabPage4 del TabControl
				}
			}
			else
			{
				// Asegúrate de que tabPage2 esté visible si la condición no se cumple
				if (!tabDatosPrincipales.TabPages.Contains(tabPage7))
				{
					tabDatosPrincipales.TabPages.Add(tabPage7); // Añade tabPage4 de nuevo al TabControl
				}
			}

		}

		public static bool RequerimientoDepartamentoTabMostrar(int idRequerimientoDepartamento, int idTipoRequerimiento, string Campo)
		{
			bool Resultado = false;
			string squery;

			if (idTipoRequerimiento > 0)
				squery = string.Format("Exec RequerimientoDepartamentoTabMostrar {0}, {1}, '{2}'", 
					idRequerimientoDepartamento, idTipoRequerimiento, Campo);
			else
				squery = string.Format("Exec RequerimientoDepartamentoTabMostrar {0}, 0, '{1}'", 
					idRequerimientoDepartamento, Campo);

			Resultado = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, squery, true);
			return Resultado;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoCheckListCrear frmrclc = new frmRequerimientoCheckListCrear(idRequerimientoDepartamento))
			{		
				if (DialogResult.OK == frmrclc.ShowDialog())
				{
				}
				RequerimientoAreaAtencionActivosObtener();
			}
		}

		private void RequerimientoAreaAtencionActivosObtener()
		{
			// Quitar la ordenación antes de actualizar los datos
			this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Bands[0].SortedColumns.Clear();
			// Cargar los nuevos datos
			string squery = String.Format("Exec RequerimientoAreaAtencionActivosObtener {0}", idRequerimientoDepartamento);
			this.ugdRequerimientoAreaAtencionActivosObtener.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			// Restaurar la selección después de cargar los datos
			RestaurarSeleccion();
			// Verificar si hay filas y aplicar la ordenación nuevamente
			if (this.ugdRequerimientoAreaAtencionActivosObtener.Rows.Count > 0)
			{
				// Volver a aplicar la ordenación a la columna "AreaAtencionTipo"
				this.ugdRequerimientoAreaAtencionActivosObtener.DisplayLayout.Bands[0].SortedColumns.Add("AreaAtencionTipo", false, true);
			}
		}

		private void RestaurarSeleccion()
		{
			// Verificar si hay datos cargados en el grid
			if (this.ugdRequerimientoAreaAtencionActivosObtener.DataSource != null)
			{
				// Iterar sobre cada fila del grid
				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.ugdRequerimientoAreaAtencionActivosObtener.Rows)
				{
					// Obtener el ID de la fila actual
					int idRequerimientoAreaAtencion = (int)row.Cells["idRequerimientoAreaAtencion"].Value;

					// Verificar si el ID está en la lista de seleccionados
					if (idsSeleccionados.Contains(idRequerimientoAreaAtencion))
					{
						// Marcar la celda de selección como verdadera
						row.Cells["Seleccionar"].Value = true;
					}
					else
					{
						// Asegurarse de que la celda de selección esté desmarcada si no está en la lista
						row.Cells["Seleccionar"].Value = false;
					}
				}
			}
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ugdRequerimientoAreaAtencionActivosObtener_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) 
		{
			try
			{
				int idTipoSeleccionado = 0;

				// Verificar si el combo tiene valor seleccionado
				if (this.cmbTipoRequerimiento.Value != null)
				{
					idTipoSeleccionado = Convert.ToInt32(this.cmbTipoRequerimiento.Value);
				}

				// Llamar al SP enviando idRequerimientoDepartamento y tipo seleccionado
				bool mostrarColumna = RequerimientoDepartamentoTabMostrar(
					idRequerimientoDepartamento, 
					idTipoSeleccionado, 
					"Asignacion"
					);

				// Ocultar la columna si no se debe mostrar
				if (!mostrarColumna)
				{
					UltraGridColumn column = e.Layout.Bands[0].Columns["Nombre"];
					if (column != null)
					{
						column.Hidden = true; // Oculta la columna
					}
				}
				else
				{
					// Asegurarse que la columna esté visible si corresponde
					UltraGridColumn column = e.Layout.Bands[0].Columns["Nombre"];
					if (column != null)
					{
						column.Hidden = false;
					}
				}
			}
			catch (Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void tabPage7_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCiudad_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoBodegaXIdCiudad();
		}

		private void cmbResponsable_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void label1_Click_1(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipoRequerimiento_ValueChanged_1(object sender, System.EventArgs e)
		{
			//Ocultar pestaña panel si no es sistemas
			OcultarPestañaPanel();
		  CargarAsignado();
			CargarEstadoRequerimiento();
			ControlNuevo();
		}

	}
}
