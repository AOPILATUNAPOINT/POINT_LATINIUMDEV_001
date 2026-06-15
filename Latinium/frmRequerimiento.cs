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
using Latinium.Services;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimiento.

	/// </summary>
	public class frmRequerimiento : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsAsignado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipo;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimiento;
		/// <summary>
		/// Variable del diseñador requerida.
		bool bVer = false;
		bool bNuevo = false;
		int iNuevo = 1;
		int iEditar = 2;
		int iEliminar = 3;
		int iVer = 4;
		//ids de permisos
		public static int iAgregar = 1;
		public static int iAsignar = 2;
		public static int iSeguimiento = 3;
		public static int iPermisos = 4;
		public static int iConfiguracion = 5;
		public static int iValoresTecnicos = 6;
		public static string sBaseRequerimeintos = "POINT";
		//PRODUCCION
		//public static string sconexionPoint = @"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='POINT';Data Source='192.168.1.165,56399'";
		//DESARROLLO
		public static string sconexionPoint = @"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='DESARROLLO';Data Source='192.168.1.165,56399'";
		//PRODUCCIÓN
		public static string sconexionInicio = @"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'";
		
		public static int iConfiguracionDocumento = 2;
		public static int iConfiguracionTamano = 1;
		public static int iConfiguracionColorNotificacion = 4;
		public static int iConfiguracionTimerbtnRequerimiento = 5;
		public static int iConfiguracionTamanoPaginacion = 6;
		public static int iConfiguracionAudioNotificacion = 7;
		public static int iConfiguracionAudioNotificacionAlerta = 8;
		public static int iConfiguracionAudioNotificacionPosicion = 9;
		public static int iConfiguracionFormularioRequerimientoPosicion = 10;
		public static int iConfiguracionAudioNotificacionPosicionSinResumen = 11;
		public static int iConfiguracionDiasMinimoFechaCreacionCOntrolMantenimiento = 13;

		public static int iRequerimientoMensajeNuevo = 1;
		public static int iRequerimientoMensajeCambioEstado = 2;
		public static int iRequerimientoMensajeAprobado = 3;

		public static int idTipoRequerimientoNuevaFuncionMejora = 3;

		public static int idUsuarioInicio = 0;
		public static int idCerrado = 7;

		//Pagina Inicial
		public static int Pagina = 1;
		public static int TotalPagina = 1;
		public static int TamanoPagina = 10;
		
		//Ids estados seguimiento
		public static int iEstadoRequerimientoNuevo = 1;
		public static int iEstadoRequerimientoAnalisis = 2;
		public static int iEstadoRequerimientoAsignado = 3;
		public static int iEstadoRequerimientoEnProceso = 4;
		public static int iEstadoRequerimientoPruebas = 5;
		public static int iEstadoRequerimientoAprobado = 6;
		public static int iEstadoRequerimientoCerrado = 7;
		public static int iEstadoRequerimientoRechazado = 8;
		public static int iEstadoRequerimientoResuelto = 13;
		public static int iEstadoRequerimientoNuevaFuncion = 21;
		
		//ids rol
		public static int iRequerimientoRolAdministrador = 1;
		public static int iRequerimientoRolTecnico = 2;
		public static int iRequerimientoRolUsuario = 3;

		//idRequerimientoDepartamento
		public static int iRequerimientoDepartamentoSistemas = 1;
		public static int iRequerimientoControlMantenimientoLocales = 2;
		public static int iRequerimientoDepartamentoMarketing = 3;
		public static int iRequerimientoAdecuacionesMantenimiento = 4;
		public static int iRequerimientoInfraestructura = 5;
		public static int iRequerimientoDePersonal = 6;
		public static int iRequerimientoConfiguracion = 9;
		//Grupos 
		public static int idGrupoAdministrador = 1;

		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblAsignacionesMias;
		private System.Windows.Forms.CheckBox chkAsignacionesMias;
		public System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRequerimiento;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		public Infragistics.Win.UltraWinGrid.UltraGrid udgSeguimiento;
		private System.Windows.Forms.Button btnPolitica;
		private System.Windows.Forms.GroupBox groupBox5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstadoRequerimientoUsuario;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgEstadoRequerimientoUsuario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentas;
		private System.Windows.Forms.Button btnNotificacion;
		private System.Windows.Forms.Label lblRegistros;
		private System.Windows.Forms.Button btnInicio;
		private System.Windows.Forms.Button btnAnterior;
		private System.Windows.Forms.Button btnFin;
		private System.Windows.Forms.Button btnSiguiente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPagina;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label lblPersonal;
		public int progress;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		public int idRequerimientoDepartamento = 0;
		public int idRequerimientoRol = 0;
		private System.Windows.Forms.Button button1;
		public string sRequerimientoRol = "";
		public int idCargo = 0;
		private System.Windows.Forms.Timer cierreTimer;

		public frmRequerimiento(int IdDepartamento)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			progress = 0;
			idRequerimientoDepartamento = IdDepartamento;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Titulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAsignacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoDias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTentativaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotaInterna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Responsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UltimaNota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaUltimaNota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioRespondeNota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imagen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntregaTec");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoDiasTec");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimiento));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimientoUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imagen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 5, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimientoUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imagen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Titulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTentativaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotaInterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UltimaNota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimaNota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioRespondeNota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imagen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntregaTec");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoDiasTec");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			this.udsRequerimiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEmpresa = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsAsignado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.button1 = new System.Windows.Forms.Button();
			this.lblRegistros = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblAsignacionesMias = new System.Windows.Forms.Label();
			this.chkAsignacionesMias = new System.Windows.Forms.CheckBox();
			this.label12 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoRequerimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.txtPagina = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnFin = new System.Windows.Forms.Button();
			this.btnAnterior = new System.Windows.Forms.Button();
			this.btnInicio = new System.Windows.Forms.Button();
			this.btnSiguiente = new System.Windows.Forms.Button();
			this.btnPolitica = new System.Windows.Forms.Button();
			this.udgEstadoRequerimientoUsuario = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsEstadoRequerimientoUsuario = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNotificacion = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.udgSeguimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.chkCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoUsuario)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).BeginInit();
			this.groupBox5.SuspendLayout();
			this.SuspendLayout();
			// 
			// udsRequerimiento
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn5.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn6.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn9.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn10.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn11.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn14.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn16.DataType = typeof(System.UInt32);
			ultraDataColumn16.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn25.DataType = typeof(int);
			this.udsRequerimiento.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26});
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
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1341, 55);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 27);
			this.btnSalir.TabIndex = 808;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(1341, 18);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(96, 28);
			this.btnNuevo.TabIndex = 1020;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nuevo");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(1096, 24);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(94, 27);
			this.btnConsultar.TabIndex = 1019;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnConsultar, "Consultar");
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// udsResponsable
			// 
			ultraDataColumn27.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn27,
																															 ultraDataColumn28});
			// 
			// udsEmpresa
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.udsEmpresa.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn29,
																													 ultraDataColumn30});
			// 
			// udsAsignado
			// 
			ultraDataColumn31.DataType = typeof(int);
			this.udsAsignado.Band.Columns.AddRange(new object[] {
																														ultraDataColumn31,
																														ultraDataColumn32});
			// 
			// udsEstado
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn33,
																													ultraDataColumn34});
			// 
			// udsRequerimientoTipo
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.udsRequerimientoTipo.Band.Columns.AddRange(new object[] {
																																		 ultraDataColumn35,
																																		 ultraDataColumn36});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-701, 175);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2925, 10);
			this.groupBox1.TabIndex = 1031;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.button1);
			this.groupBox2.Controls.Add(this.lblRegistros);
			this.groupBox2.Controls.Add(this.cmbEstado);
			this.groupBox2.Controls.Add(this.lblEstado);
			this.groupBox2.Controls.Add(this.lblAsignacionesMias);
			this.groupBox2.Controls.Add(this.chkAsignacionesMias);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.btnConsultar);
			this.groupBox2.Location = new System.Drawing.Point(19, 83);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1303, 74);
			this.groupBox2.TabIndex = 1038;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Búsqueda";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// button1
			// 
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1200, 24);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(86, 26);
			this.button1.TabIndex = 1057;
			this.button1.Text = "&Exportar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// lblRegistros
			// 
			this.lblRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistros.ForeColor = System.Drawing.Color.Red;
			this.lblRegistros.Location = new System.Drawing.Point(784, 24);
			this.lblRegistros.Name = "lblRegistros";
			this.lblRegistros.Size = new System.Drawing.Size(288, 26);
			this.lblRegistros.TabIndex = 1056;
			// 
			// cmbEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance1;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 217;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(893, 28);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(182, 25);
			this.cmbEstado.TabIndex = 1055;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(787, 28);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(125, 26);
			this.lblEstado.TabIndex = 1054;
			this.lblEstado.Text = "* Estado:";
			this.lblEstado.Visible = false;
			// 
			// lblAsignacionesMias
			// 
			this.lblAsignacionesMias.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAsignacionesMias.ForeColor = System.Drawing.Color.Red;
			this.lblAsignacionesMias.Location = new System.Drawing.Point(1085, 120);
			this.lblAsignacionesMias.Name = "lblAsignacionesMias";
			this.lblAsignacionesMias.Size = new System.Drawing.Size(230, 27);
			this.lblAsignacionesMias.TabIndex = 1053;
			// 
			// chkAsignacionesMias
			// 
			this.chkAsignacionesMias.Location = new System.Drawing.Point(970, 120);
			this.chkAsignacionesMias.Name = "chkAsignacionesMias";
			this.chkAsignacionesMias.Size = new System.Drawing.Size(96, 28);
			this.chkAsignacionesMias.TabIndex = 1052;
			this.chkAsignacionesMias.Text = "Asignaciones mías";
			this.chkAsignacionesMias.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(557, 28);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(47, 18);
			this.label12.TabIndex = 1051;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(278, 28);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 18);
			this.label17.TabIndex = 1047;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(19, 28);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 18);
			this.label18.TabIndex = 1046;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(624, 28);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(144, 24);
			this.txtNumero.TabIndex = 1050;
			this.toolTip1.SetToolTip(this.txtNumero, "Busqueda por número, responsable, asignado, título.");
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(346, 28);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(192, 24);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1045;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(77, 28);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(182, 24);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1044;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// cmbTipoRequerimiento
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRequerimiento.Appearance = appearance5;
			this.cmbTipoRequerimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoRequerimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRequerimiento.DataSource = this.udsRequerimientoTipo;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 209;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTipoRequerimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoRequerimiento.DisplayMember = "Descripcion";
			this.cmbTipoRequerimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRequerimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRequerimiento.Location = new System.Drawing.Point(950, 28);
			this.cmbTipoRequerimiento.Name = "cmbTipoRequerimiento";
			this.cmbTipoRequerimiento.Size = new System.Drawing.Size(202, 25);
			this.cmbTipoRequerimiento.TabIndex = 1049;
			this.cmbTipoRequerimiento.ValueMember = "idRequerimientoTipo";
			this.cmbTipoRequerimiento.Visible = false;
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(845, 28);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(120, 27);
			this.label11.TabIndex = 1048;
			this.label11.Text = "Requerimiento:";
			this.label11.Visible = false;
			// 
			// cmbEmpresas
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance6;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresas.DataSource = this.udsEmpresa;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 172;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(1286, 28);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(192, 25);
			this.cmbEmpresas.TabIndex = 1043;
			this.cmbEmpresas.ValueMember = "idempresa";
			this.cmbEmpresas.Visible = false;
			// 
			// cmbResponsable
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance7;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.udsResponsable;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(125, 28);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.ReadOnly = true;
			this.cmbResponsable.Size = new System.Drawing.Size(297, 25);
			this.cmbResponsable.TabIndex = 1042;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(1171, 28);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(120, 26);
			this.label4.TabIndex = 1041;
			this.label4.Text = "* Empresa:";
			this.label4.Visible = false;
			// 
			// lblPersonal
			// 
			this.lblPersonal.AutoSize = true;
			this.lblPersonal.Location = new System.Drawing.Point(10, 28);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(60, 18);
			this.lblPersonal.TabIndex = 1038;
			this.lblPersonal.Text = "* Usuario";
			// 
			// txtPagina
			// 
			this.txtPagina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtPagina.Appearance = appearance8;
			this.txtPagina.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPagina.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPagina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagina.Enabled = false;
			this.txtPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPagina.Location = new System.Drawing.Point(86, 360);
			this.txtPagina.Name = "txtPagina";
			this.txtPagina.ReadOnly = true;
			this.txtPagina.Size = new System.Drawing.Size(68, 23);
			this.txtPagina.TabIndex = 1051;
			this.toolTip1.SetToolTip(this.txtPagina, "Busqueda por número, responsable, asignado, título.");
			this.txtPagina.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// btnFin
			// 
			this.btnFin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFin.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFin.Image = ((System.Drawing.Image)(resources.GetObject("btnFin.Image")));
			this.btnFin.Location = new System.Drawing.Point(192, 360);
			this.btnFin.Name = "btnFin";
			this.btnFin.Size = new System.Drawing.Size(29, 18);
			this.btnFin.TabIndex = 1053;
			this.toolTip1.SetToolTip(this.btnFin, "Página final");
			this.btnFin.Click += new System.EventHandler(this.btnFin_Click);
			// 
			// btnAnterior
			// 
			this.btnAnterior.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnterior.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAnterior.Image = ((System.Drawing.Image)(resources.GetObject("btnAnterior.Image")));
			this.btnAnterior.Location = new System.Drawing.Point(48, 360);
			this.btnAnterior.Name = "btnAnterior";
			this.btnAnterior.Size = new System.Drawing.Size(29, 18);
			this.btnAnterior.TabIndex = 14;
			this.toolTip1.SetToolTip(this.btnAnterior, "Anterior");
			this.btnAnterior.Click += new System.EventHandler(this.btnAnterior_Click);
			// 
			// btnInicio
			// 
			this.btnInicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnInicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnInicio.Image = ((System.Drawing.Image)(resources.GetObject("btnInicio.Image")));
			this.btnInicio.Location = new System.Drawing.Point(19, 360);
			this.btnInicio.Name = "btnInicio";
			this.btnInicio.Size = new System.Drawing.Size(29, 18);
			this.btnInicio.TabIndex = 13;
			this.toolTip1.SetToolTip(this.btnInicio, "Página inicio");
			this.btnInicio.Click += new System.EventHandler(this.btnInicio_Click);
			// 
			// btnSiguiente
			// 
			this.btnSiguiente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSiguiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnSiguiente.Image = ((System.Drawing.Image)(resources.GetObject("btnSiguiente.Image")));
			this.btnSiguiente.Location = new System.Drawing.Point(163, 360);
			this.btnSiguiente.Name = "btnSiguiente";
			this.btnSiguiente.Size = new System.Drawing.Size(29, 18);
			this.btnSiguiente.TabIndex = 1052;
			this.toolTip1.SetToolTip(this.btnSiguiente, "Siguiente");
			this.btnSiguiente.Click += new System.EventHandler(this.btnSiguiente_Click);
			// 
			// btnPolitica
			// 
			this.btnPolitica.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnPolitica.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnPolitica.CausesValidation = false;
			this.btnPolitica.Image = ((System.Drawing.Image)(resources.GetObject("btnPolitica.Image")));
			this.btnPolitica.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPolitica.Location = new System.Drawing.Point(1341, 92);
			this.btnPolitica.Name = "btnPolitica";
			this.btnPolitica.Size = new System.Drawing.Size(96, 28);
			this.btnPolitica.TabIndex = 1052;
			this.btnPolitica.Text = "&Políticas";
			this.btnPolitica.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnPolitica, "Póliticas");
			this.btnPolitica.Click += new System.EventHandler(this.btnPolitica_Click);
			// 
			// udgEstadoRequerimientoUsuario
			// 
			this.udgEstadoRequerimientoUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgEstadoRequerimientoUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgEstadoRequerimientoUsuario.DataSource = this.udsEstadoRequerimientoUsuario;
			appearance9.BackColor = System.Drawing.Color.White;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Appearance = appearance9;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 95;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 56;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 24;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 27;
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 97;
			ultraGridColumn14.AutoEdit = false;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance10;
			ultraGridColumn14.Header.Caption = "C.";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 29;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.Caption = "Estado";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 115;
			ultraGridColumn16.AutoEdit = false;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.Caption = "N°";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 21;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand5.Header.Caption = "Detalle";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.udgEstadoRequerimientoUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.udgEstadoRequerimientoUsuario.DisplayLayout.GroupByBox.Hidden = true;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.udgEstadoRequerimientoUsuario.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgEstadoRequerimientoUsuario.FlatMode = true;
			this.udgEstadoRequerimientoUsuario.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgEstadoRequerimientoUsuario.Location = new System.Drawing.Point(19, 28);
			this.udgEstadoRequerimientoUsuario.Name = "udgEstadoRequerimientoUsuario";
			this.udgEstadoRequerimientoUsuario.Size = new System.Drawing.Size(192, 323);
			this.udgEstadoRequerimientoUsuario.TabIndex = 762;
			this.toolTip1.SetToolTip(this.udgEstadoRequerimientoUsuario, "Nuevo");
			this.udgEstadoRequerimientoUsuario.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.udgEstadoRequerimientoUsuario_InitializeRow);
			this.udgEstadoRequerimientoUsuario.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgEstadoRequerimientoUsuario_AfterCellUpdate);
			// 
			// udsEstadoRequerimientoUsuario
			// 
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(bool);
			ultraDataColumn42.DataType = typeof(int);
			this.udsEstadoRequerimientoUsuario.Band.Columns.AddRange(new object[] {
																																							ultraDataColumn37,
																																							ultraDataColumn38,
																																							ultraDataColumn39,
																																							ultraDataColumn40,
																																							ultraDataColumn41,
																																							ultraDataColumn42});
			// 
			// btnNotificacion
			// 
			this.btnNotificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNotificacion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNotificacion.CausesValidation = false;
			this.btnNotificacion.Image = ((System.Drawing.Image)(resources.GetObject("btnNotificacion.Image")));
			this.btnNotificacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNotificacion.Location = new System.Drawing.Point(1341, 129);
			this.btnNotificacion.Name = "btnNotificacion";
			this.btnNotificacion.Size = new System.Drawing.Size(96, 27);
			this.btnNotificacion.TabIndex = 1054;
			this.btnNotificacion.Text = "&Notf.";
			this.btnNotificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNotificacion, "Notificaciones para el usuario");
			this.btnNotificacion.Click += new System.EventHandler(this.btnNotificacion_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.lblAsignado);
			this.groupBox3.Controls.Add(this.cmbDepartamento);
			this.groupBox3.Controls.Add(this.lblPersonal);
			this.groupBox3.Controls.Add(this.cmbTipoRequerimiento);
			this.groupBox3.Controls.Add(this.label11);
			this.groupBox3.Controls.Add(this.cmbEmpresas);
			this.groupBox3.Controls.Add(this.cmbResponsable);
			this.groupBox3.Controls.Add(this.label4);
			this.groupBox3.Location = new System.Drawing.Point(19, 9);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1303, 65);
			this.groupBox3.TabIndex = 1050;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos principales";
			this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(451, 28);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(115, 26);
			this.lblAsignado.TabIndex = 1051;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// cmbDepartamento
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance18;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Descripción";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 218;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(586, 28);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.ReadOnly = true;
			this.cmbDepartamento.Size = new System.Drawing.Size(240, 25);
			this.cmbDepartamento.TabIndex = 1050;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.btnFin);
			this.groupBox4.Controls.Add(this.txtPagina);
			this.groupBox4.Controls.Add(this.btnAnterior);
			this.groupBox4.Controls.Add(this.btnInicio);
			this.groupBox4.Controls.Add(this.btnSiguiente);
			this.groupBox4.Controls.Add(this.udgSeguimiento);
			this.groupBox4.Location = new System.Drawing.Point(269, 203);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1168, 397);
			this.groupBox4.TabIndex = 1051;
			this.groupBox4.TabStop = false;
			this.groupBox4.Enter += new System.EventHandler(this.groupBox4_Enter);
			// 
			// udgSeguimiento
			// 
			this.udgSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgSeguimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgSeguimiento.DataSource = this.udsRequerimiento;
			appearance19.BackColor = System.Drawing.Color.White;
			this.udgSeguimiento.DisplayLayout.Appearance = appearance19;
			ultraGridColumn19.AutoEdit = false;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 51;
			ultraGridColumn20.AutoEdit = false;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 67;
			ultraGridColumn21.AutoEdit = false;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 58;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "Número";
			ultraGridColumn22.Header.Fixed = true;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Width = 91;
			ultraGridColumn23.AutoEdit = false;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Width = 177;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.Caption = "Título";
			ultraGridColumn24.Header.VisiblePosition = 12;
			ultraGridColumn24.Width = 196;
			ultraGridColumn25.AutoEdit = false;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn25.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn25.Header.Caption = "F. asignación";
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Width = 102;
			ultraGridColumn26.AutoEdit = false;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn26.Header.Caption = "Días";
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.Width = 34;
			ultraGridColumn27.AutoEdit = false;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn27.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn27.Header.Caption = "F. tentativa";
			ultraGridColumn27.Header.VisiblePosition = 17;
			ultraGridColumn27.Width = 109;
			ultraGridColumn28.AutoEdit = false;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn28.Header.VisiblePosition = 18;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.AutoEdit = false;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 60;
			ultraGridColumn30.Header.VisiblePosition = 20;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 40;
			ultraGridColumn31.Header.VisiblePosition = 22;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 45;
			ultraGridColumn32.Header.VisiblePosition = 23;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 90;
			ultraGridColumn33.Header.VisiblePosition = 24;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 57;
			ultraGridColumn34.Header.VisiblePosition = 26;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 61;
			ultraGridColumn35.AutoEdit = false;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn35.Header.VisiblePosition = 11;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 114;
			ultraGridColumn36.AutoEdit = false;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn36.Header.VisiblePosition = 9;
			ultraGridColumn36.Width = 162;
			ultraGridColumn37.AutoEdit = false;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Width = 143;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.Caption = "Última nota";
			ultraGridColumn38.Header.VisiblePosition = 14;
			ultraGridColumn38.Width = 206;
			ultraGridColumn39.AutoEdit = false;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn39.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn39.Header.Caption = "Fecha nota";
			ultraGridColumn39.Header.VisiblePosition = 19;
			ultraGridColumn39.Width = 102;
			ultraGridColumn40.AutoEdit = false;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn40.Header.Caption = "Usuario nota";
			ultraGridColumn40.Header.VisiblePosition = 21;
			ultraGridColumn40.Width = 131;
			ultraGridColumn41.Header.VisiblePosition = 28;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.AutoEdit = false;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn42.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn42.Header.Caption = "F. entrega técnico";
			ultraGridColumn42.Header.VisiblePosition = 25;
			ultraGridColumn42.Width = 140;
			ultraGridColumn43.AutoEdit = false;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn43.Header.Caption = "Dias técnico";
			ultraGridColumn43.Header.VisiblePosition = 27;
			ultraGridColumn43.Width = 92;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn44.Header.VisiblePosition = 8;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			appearance20.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn45.CellAppearance = appearance20;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			appearance21.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn45.CellButtonAppearance = appearance21;
			ultraGridColumn45.Header.Caption = "...";
			ultraGridColumn45.Header.Fixed = true;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 25;
			ultraGridColumn46.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn46.Header.Caption = "Estado";
			ultraGridColumn46.Header.VisiblePosition = 7;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn46.Width = 122;
			ultraGridColumn47.AutoEdit = false;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn47.Header.Caption = "N°";
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Width = 26;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47});
			ultraGridBand7.Header.Caption = "Detalle";
			ultraGridBand7.HeaderVisible = true;
			ultraGridBand7.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgSeguimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.udgSeguimiento.DisplayLayout.GroupByBox.Hidden = true;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.HeaderAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.udgSeguimiento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgSeguimiento.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgSeguimiento.Location = new System.Drawing.Point(19, 28);
			this.udgSeguimiento.Name = "udgSeguimiento";
			this.udgSeguimiento.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgSeguimiento.Size = new System.Drawing.Size(1130, 323);
			this.udgSeguimiento.TabIndex = 12;
			this.udgSeguimiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.udgSeguimiento_InitializeRow);
			this.udgSeguimiento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgSeguimiento_ClickCellButton);
			this.udgSeguimiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgSeguimiento_InitializeLayout);
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox5.Controls.Add(this.chkCuentas);
			this.groupBox5.Controls.Add(this.udgEstadoRequerimientoUsuario);
			this.groupBox5.Location = new System.Drawing.Point(19, 203);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(231, 397);
			this.groupBox5.TabIndex = 1053;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Estados de requerimiento";
			// 
			// chkCuentas
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCuentas.Appearance = appearance26;
			this.chkCuentas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuentas.Location = new System.Drawing.Point(19, 55);
			this.chkCuentas.Name = "chkCuentas";
			this.chkCuentas.Size = new System.Drawing.Size(16, 15);
			this.chkCuentas.TabIndex = 1054;
			this.chkCuentas.CheckedChanged += new System.EventHandler(this.chkCuentas_CheckedChanged);
			// 
			// frmRequerimiento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1456, 616);
			this.Controls.Add(this.btnNotificacion);
			this.Controls.Add(this.btnPolitica);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox5);
			this.Name = "frmRequerimiento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Control de requerimientos";
			this.Load += new System.EventHandler(this.frmRequerimiento_Load);
			this.Closed += new System.EventHandler(this.frmRequerimiento_Closed);
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsAsignado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipo)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRequerimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoUsuario)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).EndInit();
			this.groupBox5.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void IngresarTicketera(int idUsuario, int idTicketera)
		{
			// Llamada DIRECTA al método estático
			if (!RequerimientoAccesoService.UsuarioTieneAcceso(idUsuario, idTicketera))
			{
				MessageBox.Show("No tienes permisos para acceder a esta ticketera.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmRequerimiento_Load(object sender, System.EventArgs e)
		{
			
			//Cargar departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			if(idRequerimientoDepartamento > 0)
			{
				this.cmbDepartamento.Value = idRequerimientoDepartamento;
			}
			//Localización de la pantalla requerimiento
			int Valiorx = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMinimo", frmRequerimiento.iConfiguracionFormularioRequerimientoPosicion),false);
			int Valiory = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMaximo", frmRequerimiento.iConfiguracionFormularioRequerimientoPosicion),false);
			this.Location = new System.Drawing.Point(Valiorx, Valiory);

				string squery = "";
				try
				{
					squery = String.Format("Exec RequerimientoConfiguracionModel '{0}',{1}", "Valor", iConfiguracionTamanoPaginacion);
					TamanoPagina = Funcion.iEjecucion(sconexionPoint, squery, false);
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				}
				#region Settear paginacion
				Pagina = 1;
				TotalPagina = 1;
				#endregion Settear paginacion
				//Carga id usuario tabla Inicio
				idUsuarioInicio = Funcion.iEjecucion(sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
				//Cargar lista empresas vinculadas al usuario de ingreso
				this.cmbEmpresas.DataSource = Funcion.dsEjecucion(sconexionInicio, string.Format("Select DISTINCT(e.idEmpresa), e.Empresa From Empresas e Inner Join EmpresasUsuarios eu On eu.idEmpresa = e.idEmpresa Where eu.idUsuario = {0} Order By e.Empresa", idUsuarioInicio));
				//Cargar id de empresa de ingreso
				this.cmbEmpresas.Value = Funcion.iEjecucion(sconexionInicio, string.Format("Select Top 1 e.idEmpresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()),false);
				
				//Cargar el nombre de la empresa que se logeo
				string sempresa = Funcion.sEjecucion(sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));
				//CargarResponsable
				this.cmbResponsable.DataSource = Funcion.dsEjecucion(sconexionInicio, "Exec RequerimientoResponsable");
				try
				{
					//Seleccionar el personal que ingreso al sistema que es responsable
					this.cmbResponsable.Value = idUsuarioInicio;
					if(idUsuarioInicio == 0)
					{

						MessageBox.Show(string.Format("Su usuario {0}, no se encuentra registrado como usuario de la empresa {1}", 
							MenuLatinium.stUsuario,
							sempresa
							), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}
			
			IngresarTicketera(idUsuarioInicio, idRequerimientoDepartamento);

			string nombreCompleto = "";

			if (this.cmbResponsable.ActiveRow != null)
			{
				nombreCompleto = this.cmbResponsable.Text;
			}

//			if((int)this.cmbEmpresas.Value == 1)
//			{
//				string getsconexionEmpresaSeguimiento = Funcion.ObtenerCadenaConexionEmpresa(
//					(int)this.cmbEmpresas.Value,
//					frmRequerimiento.sconexionInicio,
//					frmSeguimiento.sconexionEmpresaSeguimiento
//					);
//
//				// Obtener idCargo desde la tabla Nomina
//				string consulting = String.Format("SELECT idPersonal FROM Personal WHERE Nombre = '{0}'", nombreCompleto);
//				int idPersonalConsulting = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, consulting, false);
//
//				string consultingNomina = String.Format("SELECT idCargo FROM Nomina WHERE idPersonal = '{0}'", idPersonalConsulting);
//				idCargo = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, consultingNomina, false);
//
//				// Validar permiso usando el procedimiento almacenado
//				bool tienePermiso = false;
//
//				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
//				using (SqlCommand cmd = new SqlCommand("RequerimientoSeguridadCargoValidar", conn))
//				{
//					cmd.CommandType = CommandType.StoredProcedure;
//
//					SqlParameter param1 = new SqlParameter("@idCargo", SqlDbType.Int);
//					param1.Value = idCargo;
//					cmd.Parameters.Add(param1);
//
//					SqlParameter param2 = new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int);
//					param2.Value = idRequerimientoDepartamento;
//					cmd.Parameters.Add(param2);
//
//					conn.Open();
//					object result = cmd.ExecuteScalar();
//					conn.Close();
//
//					tienePermiso = (result != null && Convert.ToInt32(result) == 1);
//				}
//
//
//				if (!tienePermiso)
//				{
//					MessageBox.Show("No tiene permisos para acceder a esta pantalla.", "Acceso denegado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//
//					// Crear y configurar el timer
//					cierreTimer = new System.Windows.Forms.Timer();
//					cierreTimer.Interval = 100; // milisegundos
//					cierreTimer.Tick += new EventHandler(CerrarFormulario);
//					cierreTimer.Start();
//				}
//			}
//			else if((int)this.cmbEmpresas.Value != 1 && idRequerimientoDepartamento == frmRequerimiento.iRequerimientoDePersonal)
//			{
//				MessageBox.Show("No tiene permisos para acceder a esta pantalla.", "Acceso denegado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//
//				// Crear y configurar el timer
//				cierreTimer = new System.Windows.Forms.Timer();
//				cierreTimer.Interval = 100; // milisegundos
//				cierreTimer.Tick += new EventHandler(CerrarFormulario);
//				cierreTimer.Start();
//			}

				//Cargar el tipo de requerimiento
				this.cmbTipoRequerimiento.DataSource = Funcion.dsEjecucion(sconexionPoint,string.Format("Exec RequerimientoTipoConsultar")); 
				//Cargar EstadoRequerimiento
				this.cmbEstado.DataSource = Funcion.dsEjecucion(sconexionPoint,string.Format("Exec RequerimientoEstado {0}", idRequerimientoDepartamento)); 
				//Permisos dependiendo del rol de la persona que ingresa al sistema
				RequerimientoRolPermisoCon();

				//Ubicar por default las fechas al día de hoy para no estesar la base
				#region Fechas default
				DateTime dtInicio = DateTime.Today.AddMonths(-540);
				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.dtDesde.Value = Convert.ToDateTime("14/02/2024");
				this.dtHasta.Value = DateTime.Today;
				#endregion Fechas default
				//Cargar al ingresar a la pantalla.
				RequerimientoCon();
				//Crear o actualizar el menú de estados para cada usuario
				RequerimientoEstadoUsuarioCrear();
				//Visualizar el menú estados
				RequerimientoEstadoUsuarioCon();
				//Ubicar el número de notificaciones según corresponsa al usuario 
				NotificacionBoton();
				//Ubicar el rol del usuario
				RequerimientoRolUsuario();
			
		}

		private void CerrarFormulario(object sender, EventArgs e)
		{
			cierreTimer.Stop();
			cierreTimer.Dispose();
			this.Close();
		}

		private void RequerimientoRolUsuario()
		{
			//Cual es el rol del usuario logeado
			idRequerimientoRol = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, string.Format("EXEC RequerimientoRolUsuario {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento), false);
			string squery = String.Format("Exec RequerimientoRolModel '{0}', {1}", "Subtitulo", idRequerimientoRol);
			sRequerimientoRol = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);

			//Ubicar si es administrador, técnico o usuario.
			this.lblPersonal.Text = sRequerimientoRol + ": ";

			//Mostrar solo a técnicos el reporte para indicadores.
			if(sRequerimientoRol == "Usuario")
			{
				this.button1.Visible = false;
			}
		}

		public void NotificacionBoton()
		{
			this.btnNotificacion.Text = "Notf. " + RequerimientoNotificacionCount(idUsuarioInicio, idRequerimientoDepartamento);
		}

		static public string RequerimientoNotificacionCount(int idUsuario, int idRequerimientoDepartamento)
		{
			string squery = "";
			string snotificaciones = "";
			try
			{
				if(idUsuario != 0)
				{
					squery = String.Format("EXEC RequerimientoNotificacionCount {0}, {1}", idUsuario, idRequerimientoDepartamento);
					snotificaciones = Funcion.sEjecucion(sconexionPoint, squery);
				}
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(sconexionPoint, "frmRequerimiento" + ": "+ "Control de requerimientos", string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			return snotificaciones;
		}

		static public int RequerimientoNotificacionNumero(int idUsuario)
		{
			string squery = "";
			int iRespuesta = 0;
			try
			{
				if(idUsuario != 0)
				{
					squery = String.Format("EXEC RequerimientoNotificacionNumero {0}", idUsuario);
					iRespuesta = Funcion.iEjecucion(sconexionPoint, squery, false);
				}
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(sconexionPoint, "frmRequerimiento" + ": "+ "Control de requerimientos", string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			return iRespuesta;
		}

		public void RequerimientoEstadoUsuarioCrear()
		{
			string squery = "";
			try
			{
				squery = String.Format("EXEC RequerimientoEstadoUsuarioCrear {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento);
				Funcion.EjecusionEjecutaSQL(sconexionPoint, squery, false);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
		}

		public void RequerimientoEstadoUsuarioCon()
		{
			string squery = "";
			try
			{
				squery = String.Format("EXEC RequerimientoEstadoUsuarioCon {0}, {1}", (int)this.cmbResponsable.Value, idRequerimientoDepartamento);
				this.udgEstadoRequerimientoUsuario.DataSource = FuncionesProcedimientos.dtEjecucionGeneral(sconexionPoint,squery);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			//Número de registros estado requerimiento
			this.udgEstadoRequerimientoUsuario = Funcion.ContadorFilas(this.udgEstadoRequerimientoUsuario, "Row");
		}

		private void RequerimientoRolPermisoCon()
		{
			try
			{
				//Permiso Agregar
				bool getbagregacion = Funcion.bEjecucion(sconexionPoint,string.Format("EXEC RequerimientoRolPermisoCon {0}, '{1}', {2}", 
					(int)this.cmbResponsable.Value,
					iAgregar,
					idRequerimientoDepartamento
					),false); 
				if(getbagregacion)
				{
					this.btnNuevo.Enabled = true;
				}
				else
				{
					this.btnNuevo.Enabled = false;
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
				return;
			}
			
		}

		private void RequerimientoCon()
		{
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();

			int idResponsable = 0;
			int idAsignado = 0;
			string dtDate = "";
			string dtHasta = "";
			int idTipoRequerimiento = 0;
			string Numero = "";
			bool bAsignacionesMias = false;
			int idEstado = 0;

			try { idResponsable = (int)this.cmbResponsable.Value;} 
			catch(Exception ex){string e = ex.Message;}
			try { dtDate = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd HH:mm:ss");} 
			catch(Exception ex){string e = ex.Message;}
			try { dtHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd HH:mm:ss");} 
			catch(Exception ex){string e = ex.Message;}
			try { idTipoRequerimiento = (int)this.cmbTipoRequerimiento.Value;} 
			catch(Exception ex){string e = ex.Message;}
			try { Numero = (string)this.txtNumero.Text;} 
			catch(Exception ex){string e = ex.Message;}
			try { idEstado = (int)this.cmbEstado.Value;} 
			catch(Exception ex){string e = ex.Message;}

			try 
			{ 
				if(this.chkAsignacionesMias.Checked)
				{
					bAsignacionesMias = true;
				}
			} 
			catch(Exception ex)
			{
				string e = ex.Message;
			}

			//Consulta el tamaño antes de la consulta
			iRegistrosRequerimiento(idResponsable,idAsignado ,dtDate, dtHasta , idTipoRequerimiento, Numero, bAsignacionesMias, 
				idEstado, idRequerimientoDepartamento);

			try
			{
				//Realiza la consulta por páginación
				this.udgSeguimiento.DataSource = 
					FuncionesProcedimientos.dtEjecucionGeneral(sconexionPoint,string.Format("EXEC RequerimientoCon {0},{1},'{2}','{3}',{4},'{5}',{6},{7},'{8}', {9}, {10}, {11}", 
					idResponsable,
					idAsignado,
					dtDate,
					dtHasta,
					idTipoRequerimiento,
					Numero,
					bAsignacionesMias,
					idEstado,
					MenuLatinium.stDirFacturacion,
					Pagina,
					TamanoPagina,
					idRequerimientoDepartamento
					)); 
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir durante la consulta
				MessageBox.Show("Ocurrió un error durante la consulta: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}

			//Contador de filas
			this.udgSeguimiento = Funcion.ContadorFilas(this.udgSeguimiento, "Row");
		}

		public void iRegistrosRequerimiento(int idResponsable,int idAsignado ,string dtDate,string dtHasta ,int idTipoRequerimiento,
			string Numero,bool bAsignacionesMias,int idEstado, int idRequerimientoDepartamento)
		{
			int iRegistros = 0;
			string squery = "";	
			try
			{
				squery = String.Format("EXEC RequerimientoConCount {0},{1},'{2}','{3}',{4},'{5}',{6},{7},'{8}',{9}", 
					idResponsable,
					idAsignado,
					dtDate,
					dtHasta,
					idTipoRequerimiento,
					Numero,
					bAsignacionesMias,
					idEstado,
					MenuLatinium.stDirFacturacion,
					idRequerimientoDepartamento
					);
				iRegistros = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery,false);// this.udgSeguimiento.Rows.Count;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			string sRegistros = "";
	
			if(iRegistros == 1)
			{
				sRegistros = "1 REGISTRO";
			}
			else
			{
				sRegistros = iRegistros.ToString() + " REGISTROS";
			}
			this.lblRegistros.Text = sRegistros;
			CalculoPaginas(iRegistros);
		}

		private void CalculoPaginas(int iRegistros)
		{
			TotalPagina = Convert.ToInt32(iRegistros/TamanoPagina) ;
			decimal dTotalPagina = Convert.ToDecimal(iRegistros) / Convert.ToDecimal(TamanoPagina);

			if (dTotalPagina > TotalPagina)
			{
				TotalPagina = TotalPagina + 1;
			}
			//Validar si iRegistros es igual a 0, darle valor a 1 a total de la página
			if (iRegistros == 0)
			{
				TotalPagina = 1;
				Pagina = 1;
			}

			MostrarPaginaActual();
		}


		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Validar politicas
			Resultado resultado = new Resultado();
			resultado = frmRequerimiento.RequerimientoPoliticaVal(4,1, 0, (int)this.cmbResponsable.Value, idRequerimientoDepartamento);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			//Validación 5 requerimientos máximo
			resultado = frmRequerimiento.RequerimientoPoliticaVal(9,1, 0, (int)this.cmbResponsable.Value, idRequerimientoDepartamento);
			if(resultado.bResultado == false)
			{
				MessageBox.Show(string.Format(resultado.sResultado), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Validar politicas
			using (frmSeguimiento frmS = new frmSeguimiento((int)this.cmbResponsable.Value,0,(int)this.cmbEmpresas.Value,iNuevo, idRequerimientoDepartamento))
			{		
				if (DialogResult.OK == frmS.ShowDialog())
				{
				}
				RequerimientoCon(); 
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			#region Settear paginacion
			Pagina = 1;
			TotalPagina = 1;
			#endregion Settear paginacion
			#region valida Campos Generales 
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione el responsable ")) return;
			#endregion valida Campos Generales 
			RequerimientoCon();
			//Visualizar el menú estados
			RequerimientoEstadoUsuarioCon();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void udgSeguimiento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
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
			int idRequerimiento = 0;
			string Numero = "";
			string Imagen = "";

			if (e.Cell.Column.ToString() == "btnSeguimientoAnadir")	
			{
				try
				{
					idResponsable = (int)e.Cell.Row.Cells["idResponsable"].Value;
					idRequerimientoTipo = (int)e.Cell.Row.Cells["idRequerimientoTipo"].Value;
					idEmpresa = (int)e.Cell.Row.Cells["idEmpresa"].Value;
					idAsignado = (int)e.Cell.Row.Cells["idAsignado"].Value;
					Titulo = (string)e.Cell.Row.Cells["Titulo"].Value;
					Descripcion = (string)e.Cell.Row.Cells["Descripcion"].Value;
					FechaAsignacion = (DateTime)e.Cell.Row.Cells["FechaAsignacion"].Value;
					FechaTentativaEntrega = (DateTime)e.Cell.Row.Cells["FechaTentativaEntrega"].Value;
					NotaInterna = (string)e.Cell.Row.Cells["NotaInterna"].Value;
					idEstado = (int)e.Cell.Row.Cells["idEstado"].Value;
					idPantalla = (int)e.Cell.Row.Cells["idPantalla"].Value;
					idSeguimiento = (int)e.Cell.Row.Cells["idSeguimiento"].Value;
					idRequerimiento = (int)e.Cell.Row.Cells["idRequerimiento"].Value;
					Numero = (string)e.Cell.Row.Cells["Numero"].Value;

					using (frmSeguimientoAnadir frmS = new frmSeguimientoAnadir(
									 idResponsable,
									 idRequerimientoTipo,
									 idEmpresa,
									 idAsignado,
									 Titulo,
									 Descripcion,
									 FechaAsignacion,
									 FechaTentativaEntrega,
									 NotaInterna,
									 idEstado,
									 idPantalla,
									 idSeguimiento,
									 idRequerimiento,
									 Numero,
									 idRequerimientoDepartamento))
					{		
						if (DialogResult.OK == frmS.ShowDialog())
						{
					 
						}
						RequerimientoCon();
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					return;
				}
			}
		}

		private void label11_Click(object sender, System.EventArgs e)
		{
		
		}

		private void checkBox1_CheckedChanged(object sender, System.EventArgs e)
		{
			if(this.chkAsignacionesMias.Checked)
			{
				RequerimientoCon();
				//this.btnConsultar.Enabled = false;
				int count = this.udgSeguimiento.Rows.Count;
				if(count == 0)
				{
					this.lblAsignacionesMias.Text ="SIN ASIGNACIONES";
				}
			}
			else
			{
				this.lblAsignacionesMias.Text = "";
				RequerimientoCon();
			}

		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		//Consumo de políticas
		public static Resultado RequerimientoPoliticaVal(int NumeroPolitica, int idEstado , int idRequerimiento, int idPersonal, int getidRequerimientoDepartamento)
		{
			bool bResultado = false;
			string sResultado = "";
			string squery = "";
			DataSet dsResultado = new DataSet();
			DataView dvResultado= new DataView();
			Resultado resultado = new Resultado();

			squery = string.Format("EXEC RequerimientoPoliticaVal {0},{1},{2}, {3}, {4}", 
				NumeroPolitica,
				idEstado,
				idRequerimiento,
				idPersonal,
				getidRequerimientoDepartamento
				);

			dvResultado = Funcion.dvProcedimiento(frmRequerimiento.sconexionPoint, squery);
			dsResultado = dvResultado.Table.DataSet;
			bResultado = bool.Parse(dsResultado.Tables[0].Rows[0]["bResultado"].ToString());
			sResultado = dsResultado.Tables[0].Rows[0]["sResultado"].ToString();

			resultado.bResultado = bResultado;
			resultado.sResultado = sResultado;

			return resultado;
		}

		//Consumo de políticas mensaje
		public static string RequerimientoPoliticaValMensaje(int NumeroPolitica, bool Resultado)
		{
			string sSalida = "";
			if(Resultado == false)
			{
				//Permiso Agregar
				sSalida = Funcion.sEjecucion(sconexionPoint,string.Format("Select Mensaje from RequerimientoPolitica where NumeroPolitica = {0}", NumeroPolitica
					)); 
			}
			return sSalida;
		}

		private void groupBox3_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnPolitica_Click(object sender, System.EventArgs e)
		{
			//			// maximum indica el valor máximo de la barra
			//			pBar.Maximum = 100;
			//			//minimum indica el valor mínimo de la barra
			//			pBar.Minimum=0;
			//			//value indica desde donde se va a comenzar a llenar la barra, la nuestra iniciara desde cero
			//			pBar.Value=0;
			//			//Por ejemplo podemos hacer que la barra inicie desde la mitad
			//			//la siguiente instrucción indica que inicie cargando desde la mitad del tamaño de la barra
			//			//progressBar1.Value = progressBar1.Maximum / 2;
			//			//step indica el paso de la barra, entre más pequeño sea la barra tardará más en cargar
			//			pBar.Step=1;
			//           
			//			//el ciclo for cargará la barra
			//			for (int i = pBar.Minimum; i < pBar.Maximum; i=i+pBar.Step)
			//			{
			//				//esta instrucción avanza la posición actual de la barra
			//				pBar.PerformStep();
			//				//cambio de colores
			//				if (pBar.Value>0 && pBar.Value<33)
			//				{
			//					pBar.Foreground = new SolidColorBrush(Colors.Red);
			//				}
			//				else if (pBar.Value > 33 && pBar.Value < 75)
			//				{
			//					pBar.Foreground = new SolidColorBrush(Colors.Yellow);
			//				}
			//				else if (pBar.Value > 75 && pBar.Value < 100)
			//				{
			//					pBar.Foreground = new SolidColorBrush(Colors.Green);
			//				}
			//			}
			//			pBar.Value = 0;

			string spoliticas = Funcion.sEjecucion(sconexionPoint,"EXEC RequerimientoPoliticaCon"); 
			MessageBox.Show(string.Format(spoliticas), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void udgSeguimiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			int getXidRequerimientoDepartamento = 0;
			try
			{
				getXidRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			if(getXidRequerimientoDepartamento>0)
			{
				// Verificar el valor devuelto y ocultar la columna si es necesario
				if (!frmSeguimiento.RequerimientoDepartamentoTabMostrar(getXidRequerimientoDepartamento, 0, "bDescripcion"))
				{
					// Encuentra la columna por nombre
					UltraGridColumn column = e.Layout.Bands[0].Columns["Titulo"];
					if (column != null)
					{
						column.Hidden = true; // Oculta la columna
					}
				}
			}
		}

		private void udgSeguimiento_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
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

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
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

		private void udgEstadoRequerimientoUsuario_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

		}

		private void RequerimientoEstadoUsuarioSel(int idEstadoRequerimientoUsuario, bool Seleccionar)
		{
			string squery = "";
			try
			{
				squery = string.Format("Exec RequerimientoEstadoUsuarioSel {0}, {1}", idEstadoRequerimientoUsuario,Seleccionar);
				Funcion.EjecusionEjecutaSQL(sconexionPoint, squery, false);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
		}

		private void udgEstadoRequerimientoUsuario_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.Header.Caption.ToString() == sColumna)
			{
				RequerimientoEstadoUsuarioSel((int)e.Cell.Row.Cells["idEstadoRequerimientoUsuario"].Value,(bool)e.Cell.Row.Cells["Seleccionar"].Value);

			}
		}

		private void udgEstadoRequerimientoUsuario_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
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

		private void chkCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgEstadoRequerimientoUsuario.Rows.All)
				dr.Cells["Seleccionar"].Value = this.chkCuentas.Checked;
		}

		static public bool RequerimientoNotificacionCrear(int idResponsable, int idAsignado, int idRequerimientoMensaje, int getidSeguimiento, int idEstado, int getidRequerimientoDepartamento)
		{
			string squery = "";
			string Notificacion = "";
			bool bRespuesta = true;

			try
			{

				int getidRequerimiento = Funcion.iEjecucion(sconexionPoint, String.Format("Exec SeguimientoModel {0}, {1}", "idRequerimiento", getidSeguimiento), false);
				//Llamar a los mensajes
				Notificacion = RequerimientoMensajeCon(idRequerimientoMensaje, getidRequerimiento);
				//Llamar a la notificación
				squery = string.Format("Exec RequerimientoNotificacionCrear {0}, {1}, '{2}', {3}, {4},  {5}",
					idResponsable, idAsignado, Notificacion, getidRequerimiento, getidSeguimiento, getidRequerimientoDepartamento
					);
				Funcion.EjecusionEjecutaSQL(sconexionPoint, squery, false);

			}	
			catch(Exception ex)
			{
				bRespuesta = false;
				Funcion.LogSistema(sconexionPoint, "frmRequerimiento" + ": "+ "Control de requerimientos", string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			return bRespuesta;
		}

		static public string RequerimientoMensajeCon(int idRequerimientoMensaje, int getidRequerimiento)
		{
			string squery = "";
			string sRespuesta= "";
			try
			{
				//Consultar el caso
				string Numero = Funcion.sEjecucion(sconexionPoint, String.Format("Exec RequerimientoModel {0}, {1}", "Numero", getidRequerimiento));
				string Estado = Funcion.sEjecucion(sconexionPoint, String.Format("Exec RequerimientoModel {0}, {1}", "Estado", getidRequerimiento));
				//Consultar el estado del requerimiento
				//Llamar a la notificación
				squery = string.Format("Exec RequerimientoMensajeCon {0}", idRequerimientoMensaje);
				sRespuesta = Funcion.sEjecucion(sconexionPoint, squery);
				sRespuesta = string.Format(sRespuesta, Numero, Estado);
			}	
			catch(Exception ex)
			{
				sRespuesta= "";
				Funcion.LogSistema(sconexionPoint, "frmRequerimiento" + ": "+ "Control de requerimientos", string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
			return sRespuesta;
		}

		private void btnNotificacion_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoNotificacion frmrn = new frmRequerimientoNotificacion(idRequerimientoDepartamento))
			{		
				if (DialogResult.OK == frmrn.ShowDialog())
				{
				}
				RequerimientoCon(); 
				//Ubicar el número de notificaciones según corresponsa al usuario 
				NotificacionBoton();
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.ValidaLetraDigito(sender,e, '-');
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox4_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void MostrarPaginaActual()
		{
			this.txtPagina.Text = Pagina.ToString() + "/" + TotalPagina.ToString();
		}

		private void btnInicio_Click(object sender, System.EventArgs e)
		{
			Pagina = 1;
			MostrarPaginaActual();
			RequerimientoCon();
		}

		private void btnAnterior_Click(object sender, System.EventArgs e)
		{
			if(Pagina > 1)
			{
				Pagina = Pagina - 1;
			}
			MostrarPaginaActual();
			RequerimientoCon();
		}

		private void btnSiguiente_Click(object sender, System.EventArgs e)
		{
			if(Pagina < TotalPagina)
			{
				Pagina = Pagina + 1;
			}
			MostrarPaginaActual();
			RequerimientoCon();
		}

		private void btnFin_Click(object sender, System.EventArgs e)
		{
			Pagina = TotalPagina;
			MostrarPaginaActual();
			RequerimientoCon();
		}

		private void frmRequerimiento_Closed(object sender, System.EventArgs e)
		{
		}

		private void axWindowsMediaPlayer1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			using (frmrequerimientoreporte frmS = new frmrequerimientoreporte(idRequerimientoDepartamento))
			{		
				if (DialogResult.OK == frmS.ShowDialog())
				{
				}
			}
		}
	}

	public class Resultado
	{
		public bool bResultado;
		public string sResultado;
	}

}
