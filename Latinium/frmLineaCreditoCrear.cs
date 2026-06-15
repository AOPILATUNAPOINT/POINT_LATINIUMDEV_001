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
	/// Descripción breve de frmLineaCreditoCrear.
	/// </summary>
	public class frmLineaCreditoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btSeguridad;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		int idResponsable = 0;
		int idAsignado = 0;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		int idEmpresa = 0;
		bool bVer = false;
		bool bNuevo = false;
		static public int iNuevo = 1;
		static public int iAnadir = 2;
		static public int iEliminar = 3;
		static public int iVer = 4;
		bool bRequerimientosDocumento = false;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.TabControl tabRequerimiento;
		string Title = "";
		public static string sconexionEmpresaSeguimiento = @"Password=Bl45o6$9;Persist Security Info=True;User ID= {0};Initial Catalog={1};Data Source={2}";
		public 
		string getsconexionEmpresaSeguimiento = "";
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsGarantia;
		int icontadoringreso = 0;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdBienInmueble;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDestino;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGarantiasProcentajes;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTasa;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRenovacion;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaAprobacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udgBanco;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDestino;
		static public int idLineaCreditoTransaccion = 0;
		public System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMontoTotal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsLineaCreditoDestinoDoc;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsGarantiaLineaCreditoAvaluoDoc;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdLineaCreditoDestinoDoc;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdGarantiaLineaCreditoAvaluoDoc;
		int idLineaCredito = 0;
		public System.Windows.Forms.Label label6;
		public System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMontoUtilizado;
		private System.Windows.Forms.GroupBox groupBox3;
		public System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProveedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsProveedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		public System.Windows.Forms.Label lblBanco;
		public System.Windows.Forms.Label lblProveedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipo;
		int iTipo = 0;
		
		public frmLineaCreditoCrear(int ITipo, int IdLineaCredito)
		{
			iTipo = ITipo;
			//Cargar id si es necesario, o se debe realizar un mantenimiento del mismo en nuevo es 0
			idLineaCredito = IdLineaCredito;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLineaCreditoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestinoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 3, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestinoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGarantiaLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idListadoBienInmueble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BienInmueble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGarantiaLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idListadoBienInmueble");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BienInmueble");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestinoDoc", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNume");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEdicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestinoDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNume");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEdicion");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGarantiaLineaCreditoAvaluoDoc", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGarantiaLineaCreditoAvaluo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNume");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEdicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGarantiaLineaCreditoAvaluoDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGarantiaLineaCreditoAvaluo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNume");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEdicion");
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.tabRequerimiento = new System.Windows.Forms.TabControl();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.ugdDestino = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDestino = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSeguridad = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.ugdBienInmueble = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsGarantia = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbProveedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsProveedor = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblProveedor = new System.Windows.Forms.Label();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.txtMontoUtilizado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtMontoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtGarantiasProcentajes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTasa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaRenovacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFechaAprobacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udgBanco = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.ugdLineaCreditoDestinoDoc = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsLineaCreditoDestinoDoc = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugdGarantiaLineaCreditoAvaluoDoc = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsGarantiaLineaCreditoAvaluoDoc = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			this.tabRequerimiento.SuspendLayout();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDestino)).BeginInit();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdBienInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoUtilizado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantiasProcentajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRenovacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAprobacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdLineaCreditoDestinoDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCreditoDestinoDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdGarantiaLineaCreditoAvaluoDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGarantiaLineaCreditoAvaluoDoc)).BeginInit();
			this.SuspendLayout();
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(624, 16);
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
			this.tabRequerimiento.Controls.Add(this.tabPage2);
			this.tabRequerimiento.Controls.Add(this.tabPage1);
			this.tabRequerimiento.Location = new System.Drawing.Point(16, 256);
			this.tabRequerimiento.Name = "tabRequerimiento";
			this.tabRequerimiento.SelectedIndex = 0;
			this.tabRequerimiento.Size = new System.Drawing.Size(680, 280);
			this.tabRequerimiento.TabIndex = 806;
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.ugdDestino);
			this.tabPage2.Controls.Add(this.btSeguridad);
			this.tabPage2.Controls.Add(this.btCancelar);
			this.tabPage2.Controls.Add(this.btAceptar);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(672, 254);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "* Destino";
			// 
			// ugdDestino
			// 
			this.ugdDestino.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDestino.DataSource = this.udsDestino;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugdDestino.DisplayLayout.Appearance = appearance1;
			this.ugdDestino.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 84;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 168;
			ultraGridColumn4.AutoEdit = false;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Monto ($)";
			ultraGridColumn4.Header.VisiblePosition = 8;
			ultraGridColumn4.Width = 173;
			ultraGridColumn5.AutoEdit = false;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "";
			ultraGridColumn5.Header.Caption = "Tiempo (meses)";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Width = 182;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 202;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn7.CellAppearance = appearance4;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			ultraGridColumn7.CellButtonAppearance = appearance5;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn7.Width = 26;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			ultraGridColumn8.CellAppearance = appearance6;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			ultraGridColumn8.CellButtonAppearance = appearance7;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn8.Width = 18;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn9.CellAppearance = appearance8;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn9.CellButtonAppearance = appearance9;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 18;
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
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.ugdDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugdDestino.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdDestino.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDestino.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDestino.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ugdDestino.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDestino.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDestino.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDestino.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ugdDestino.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDestino.Location = new System.Drawing.Point(16, 16);
			this.ugdDestino.Name = "ugdDestino";
			this.ugdDestino.Size = new System.Drawing.Size(640, 208);
			this.ugdDestino.TabIndex = 811;
			this.ugdDestino.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDestino_ClickCellButton);
			this.ugdDestino.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdMenu_InitializeLayout);
			// 
			// udsDestino
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn4.DataType = typeof(System.UInt32);
			ultraDataColumn5.DataType = typeof(System.UInt32);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(int);
			this.udsDestino.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn3,
																													 ultraDataColumn4,
																													 ultraDataColumn5,
																													 ultraDataColumn6,
																													 ultraDataColumn7,
																													 ultraDataColumn8});
			// 
			// btSeguridad
			// 
			this.btSeguridad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSeguridad.Location = new System.Drawing.Point(8, 336);
			this.btSeguridad.Name = "btSeguridad";
			this.btSeguridad.Size = new System.Drawing.Size(88, 24);
			this.btSeguridad.TabIndex = 14;
			this.btSeguridad.Text = "Con&figuracion";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(80, 368);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 13;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 368);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 12;
			this.btAceptar.Text = "&Aceptar";
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.ugdBienInmueble);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(672, 254);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "* Garantías";
			// 
			// ugdBienInmueble
			// 
			this.ugdBienInmueble.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdBienInmueble.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdBienInmueble.DataSource = this.udsGarantia;
			appearance15.BackColor = System.Drawing.Color.White;
			this.ugdBienInmueble.DisplayLayout.Appearance = appearance15;
			this.ugdBienInmueble.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 388;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 243;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 133;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 169;
			ultraGridColumn14.Header.Caption = "Bien inmueble";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 531;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			ultraGridColumn15.CellAppearance = appearance16;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			ultraGridColumn15.CellButtonAppearance = appearance17;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 24;
			appearance18.Image = ((object)(resources.GetObject("appearance18.Image")));
			ultraGridColumn16.CellAppearance = appearance18;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			ultraGridColumn16.CellButtonAppearance = appearance19;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 21;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			ultraGridColumn17.CellAppearance = appearance20;
			appearance21.Image = ((object)(resources.GetObject("appearance21.Image")));
			ultraGridColumn17.CellButtonAppearance = appearance21;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 21;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.ugdBienInmueble.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdBienInmueble.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdBienInmueble.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdBienInmueble.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdBienInmueble.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ugdBienInmueble.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdBienInmueble.DisplayLayout.Override.HeaderAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdBienInmueble.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdBienInmueble.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ugdBienInmueble.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdBienInmueble.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdBienInmueble.Location = new System.Drawing.Point(30, 16);
			this.ugdBienInmueble.Name = "ugdBienInmueble";
			this.ugdBienInmueble.Size = new System.Drawing.Size(618, 224);
			this.ugdBienInmueble.TabIndex = 812;
			this.ugdBienInmueble.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ugdBienInmueble.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdBienInmueble_InitializeLayout);
			// 
			// udsGarantia
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn10.DataType = typeof(System.UInt32);
			ultraDataColumn11.DataType = typeof(System.UInt32);
			ultraDataColumn12.DataType = typeof(System.UInt32);
			this.udsGarantia.Band.Columns.AddRange(new object[] {
																														ultraDataColumn9,
																														ultraDataColumn10,
																														ultraDataColumn11,
																														ultraDataColumn12,
																														ultraDataColumn13});
			this.udsGarantia.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.udsGarantiaLinea_CellDataRequested);
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
			this.btnSalir.Location = new System.Drawing.Point(624, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 807;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.cmbProveedor);
			this.groupBox2.Controls.Add(this.lblProveedor);
			this.groupBox2.Controls.Add(this.cmbTipo);
			this.groupBox2.Controls.Add(this.label8);
			this.groupBox2.Controls.Add(this.txtMontoUtilizado);
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Controls.Add(this.txtSaldo);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.txtMontoTotal);
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.txtNumero);
			this.groupBox2.Controls.Add(this.txtGarantiasProcentajes);
			this.groupBox2.Controls.Add(this.txtTasa);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.dtFechaRenovacion);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.dtFechaAprobacion);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.cmbBanco);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.lblBanco);
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(600, 216);
			this.groupBox2.TabIndex = 1033;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// cmbProveedor
			// 
			this.cmbProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProveedor.Appearance = appearance26;
			this.cmbProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbProveedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProveedor.DataSource = this.udsProveedor;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 161;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbProveedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProveedor.DisplayMember = "Nombre";
			this.cmbProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProveedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProveedor.Location = new System.Drawing.Point(424, 56);
			this.cmbProveedor.Name = "cmbProveedor";
			this.cmbProveedor.Size = new System.Drawing.Size(152, 21);
			this.cmbProveedor.TabIndex = 1020;
			this.cmbProveedor.ValueMember = "idCliente";
			// 
			// udsProveedor
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn14.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udsProveedor.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn14,
																														 ultraDataColumn15});
			// 
			// lblProveedor
			// 
			this.lblProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.lblProveedor.Location = new System.Drawing.Point(296, 56);
			this.lblProveedor.Name = "lblProveedor";
			this.lblProveedor.Size = new System.Drawing.Size(100, 16);
			this.lblProveedor.TabIndex = 1019;
			this.lblProveedor.Text = "* Proveedor:";
			// 
			// cmbTipo
			// 
			this.cmbTipo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance27;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.udsTipo;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.Caption = "Descripción";
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 158;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipo.DisplayMember = "Descripcion";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(136, 56);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(152, 21);
			this.cmbTipo.TabIndex = 1018;
			this.cmbTipo.ValueMember = "idLineaCreditoTipo";
			this.cmbTipo.ValueChanged += new System.EventHandler(this.ultraCombo1_ValueChanged);
			// 
			// udsTipo
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn16.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udsTipo.Band.Columns.AddRange(new object[] {
																												ultraDataColumn16,
																												ultraDataColumn17});
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label8.Location = new System.Drawing.Point(16, 56);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(100, 24);
			this.label8.TabIndex = 1017;
			this.label8.Text = "* Tipo:";
			// 
			// txtMontoUtilizado
			// 
			this.txtMontoUtilizado.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMontoUtilizado.Appearance = appearance28;
			this.txtMontoUtilizado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMontoUtilizado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMontoUtilizado.Enabled = false;
			this.txtMontoUtilizado.FormatString = "#,##0.00";
			this.txtMontoUtilizado.Location = new System.Drawing.Point(424, 152);
			this.txtMontoUtilizado.Name = "txtMontoUtilizado";
			this.txtMontoUtilizado.NullText = "0.00";
			this.txtMontoUtilizado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMontoUtilizado.PromptChar = ' ';
			this.txtMontoUtilizado.Size = new System.Drawing.Size(152, 19);
			this.txtMontoUtilizado.TabIndex = 1016;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(304, 152);
			this.label7.Name = "label7";
			this.label7.TabIndex = 1015;
			this.label7.Text = "Monto utilizado ($):";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldo
			// 
			this.txtSaldo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance29;
			this.txtSaldo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(136, 184);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NullText = "0.00";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(152, 19);
			this.txtSaldo.TabIndex = 1014;
			this.txtSaldo.ValueChanged += new System.EventHandler(this.ultraNumericEditor1_ValueChanged);
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 184);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(53, 16);
			this.label6.TabIndex = 1013;
			this.label6.Text = "Saldo ($):";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label6.Click += new System.EventHandler(this.label6_Click);
			// 
			// txtMontoTotal
			// 
			this.txtMontoTotal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMontoTotal.Appearance = appearance30;
			this.txtMontoTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMontoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMontoTotal.FormatString = "#,##0.00";
			this.txtMontoTotal.Location = new System.Drawing.Point(136, 152);
			this.txtMontoTotal.Name = "txtMontoTotal";
			this.txtMontoTotal.NullText = "0.00";
			this.txtMontoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMontoTotal.PromptChar = ' ';
			this.txtMontoTotal.Size = new System.Drawing.Size(152, 19);
			this.txtMontoTotal.TabIndex = 1012;
			this.txtMontoTotal.ValueChanged += new System.EventHandler(this.txtMontoTotal_ValueChanged);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 152);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(88, 16);
			this.label5.TabIndex = 1011;
			this.label5.Text = "* Monto total ($):";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(48, 16);
			this.label12.TabIndex = 1009;
			this.label12.Text = "Número:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Click += new System.EventHandler(this.label12_Click);
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance31;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(136, 24);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(152, 21);
			this.txtNumero.TabIndex = 1008;
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtGarantiasProcentajes
			// 
			this.txtGarantiasProcentajes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGarantiasProcentajes.Appearance = appearance32;
			this.txtGarantiasProcentajes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGarantiasProcentajes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGarantiasProcentajes.FormatString = "#,##0.00";
			this.txtGarantiasProcentajes.Location = new System.Drawing.Point(424, 120);
			this.txtGarantiasProcentajes.Name = "txtGarantiasProcentajes";
			this.txtGarantiasProcentajes.NullText = "0.00";
			this.txtGarantiasProcentajes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGarantiasProcentajes.PromptChar = ' ';
			this.txtGarantiasProcentajes.Size = new System.Drawing.Size(152, 19);
			this.txtGarantiasProcentajes.TabIndex = 1007;
			// 
			// txtTasa
			// 
			this.txtTasa.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTasa.Appearance = appearance33;
			this.txtTasa.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTasa.FormatString = "#,##0.00";
			this.txtTasa.Location = new System.Drawing.Point(136, 120);
			this.txtTasa.Name = "txtTasa";
			this.txtTasa.NullText = "0.00";
			this.txtTasa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTasa.PromptChar = ' ';
			this.txtTasa.Size = new System.Drawing.Size(152, 19);
			this.txtTasa.TabIndex = 1006;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(296, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(120, 16);
			this.label2.TabIndex = 1005;
			this.label2.Text = "* Fecha renovación:";
			// 
			// dtFechaRenovacion
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRenovacion.Appearance = appearance34;
			this.dtFechaRenovacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaRenovacion.DateButtons.Add(dateButton1);
			this.dtFechaRenovacion.Format = "dd/MM/yyyy hh:mm:ss";
			this.dtFechaRenovacion.Location = new System.Drawing.Point(424, 88);
			this.dtFechaRenovacion.Name = "dtFechaRenovacion";
			this.dtFechaRenovacion.NonAutoSizeHeight = 23;
			this.dtFechaRenovacion.Size = new System.Drawing.Size(152, 21);
			this.dtFechaRenovacion.SpinButtonsVisible = true;
			this.dtFechaRenovacion.TabIndex = 1004;
			this.dtFechaRenovacion.Value = ((object)(resources.GetObject("dtFechaRenovacion.Value")));
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 88);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 16);
			this.label3.TabIndex = 1003;
			this.label3.Text = "* Fecha aprobación:";
			// 
			// dtFechaAprobacion
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAprobacion.Appearance = appearance35;
			this.dtFechaAprobacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAprobacion.DateButtons.Add(dateButton2);
			this.dtFechaAprobacion.Format = "dd/MM/yyyy hh:mm:ss";
			this.dtFechaAprobacion.Location = new System.Drawing.Point(136, 88);
			this.dtFechaAprobacion.Name = "dtFechaAprobacion";
			this.dtFechaAprobacion.NonAutoSizeHeight = 23;
			this.dtFechaAprobacion.Size = new System.Drawing.Size(152, 21);
			this.dtFechaAprobacion.SpinButtonsVisible = true;
			this.dtFechaAprobacion.TabIndex = 1002;
			this.dtFechaAprobacion.Value = ((object)(resources.GetObject("dtFechaAprobacion.Value")));
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(296, 120);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(136, 48);
			this.label11.TabIndex = 818;
			this.label11.Text = "* Porcentaje garantía (%):";
			// 
			// cmbBanco
			// 
			this.cmbBanco.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance36;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.udgBanco;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 157;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(424, 56);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(152, 21);
			this.cmbBanco.TabIndex = 817;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// udgBanco
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn18.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udgBanco.Band.Columns.AddRange(new object[] {
																												 ultraDataColumn18,
																												 ultraDataColumn19});
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.label4.Location = new System.Drawing.Point(16, 120);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(120, 24);
			this.label4.TabIndex = 814;
			this.label4.Text = "* Tasa (%):";
			// 
			// lblBanco
			// 
			this.lblBanco.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.lblBanco.Location = new System.Drawing.Point(296, 56);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(100, 16);
			this.lblBanco.TabIndex = 812;
			this.lblBanco.Text = "* Banco:";
			// 
			// ugdLineaCreditoDestinoDoc
			// 
			this.ugdLineaCreditoDestinoDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdLineaCreditoDestinoDoc.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdLineaCreditoDestinoDoc.DataSource = this.udsLineaCreditoDestinoDoc;
			appearance37.BackColor = System.Drawing.Color.White;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Appearance = appearance37;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 8;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Width = 17;
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn28.Width = 8;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 4;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 169;
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn31.Width = 17;
			ultraGridColumn32.Header.VisiblePosition = 11;
			ultraGridColumn32.Width = 8;
			ultraGridColumn33.Header.VisiblePosition = 12;
			ultraGridColumn33.Width = 8;
			ultraGridColumn34.Header.VisiblePosition = 13;
			ultraGridColumn34.Width = 8;
			ultraGridColumn35.Header.VisiblePosition = 14;
			ultraGridColumn35.Width = 8;
			ultraGridColumn36.Header.VisiblePosition = 15;
			ultraGridColumn36.Width = 17;
			appearance38.Image = ((object)(resources.GetObject("appearance38.Image")));
			ultraGridColumn37.CellAppearance = appearance38;
			ultraGridColumn37.Header.Caption = "...";
			ultraGridColumn37.Header.VisiblePosition = 7;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn37.Width = 8;
			appearance39.Image = ((object)(resources.GetObject("appearance39.Image")));
			ultraGridColumn38.CellAppearance = appearance39;
			ultraGridColumn38.Header.Caption = "...";
			ultraGridColumn38.Header.VisiblePosition = 8;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn38.Width = 8;
			appearance40.Image = ((object)(resources.GetObject("appearance40.Image")));
			ultraGridColumn39.CellAppearance = appearance40;
			ultraGridColumn39.Header.Caption = "...";
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn39.Width = 8;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn39});
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance41.BackColor = System.Drawing.Color.Transparent;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.CardAreaAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 10F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.HeaderAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.RowSelectorAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.Override.SelectedRowAppearance = appearance44;
			this.ugdLineaCreditoDestinoDoc.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdLineaCreditoDestinoDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdLineaCreditoDestinoDoc.Location = new System.Drawing.Point(632, 96);
			this.ugdLineaCreditoDestinoDoc.Name = "ugdLineaCreditoDestinoDoc";
			this.ugdLineaCreditoDestinoDoc.Size = new System.Drawing.Size(0, 136);
			this.ugdLineaCreditoDestinoDoc.TabIndex = 1036;
			this.ugdLineaCreditoDestinoDoc.Visible = false;
			// 
			// udsLineaCreditoDestinoDoc
			// 
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn20.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn21.DataType = typeof(System.UInt32);
			ultraDataColumn23.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(bool);
			ultraDataColumn32.DataType = typeof(System.DateTime);
			this.udsLineaCreditoDestinoDoc.Band.Columns.AddRange(new object[] {
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
			// ugdGarantiaLineaCreditoAvaluoDoc
			// 
			this.ugdGarantiaLineaCreditoAvaluoDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdGarantiaLineaCreditoAvaluoDoc.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DataSource = this.udsGarantiaLineaCreditoAvaluoDoc;
			appearance45.BackColor = System.Drawing.Color.White;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Appearance = appearance45;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Width = 8;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 2;
			ultraGridColumn42.Width = 8;
			ultraGridColumn43.Header.VisiblePosition = 3;
			ultraGridColumn43.Width = 17;
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 6;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 4;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 169;
			ultraGridColumn47.Header.VisiblePosition = 9;
			ultraGridColumn47.Width = 17;
			ultraGridColumn48.Header.VisiblePosition = 11;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 12;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 13;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 14;
			ultraGridColumn51.Width = 8;
			ultraGridColumn52.Header.VisiblePosition = 15;
			ultraGridColumn52.Width = 17;
			appearance46.Image = ((object)(resources.GetObject("appearance46.Image")));
			ultraGridColumn53.CellAppearance = appearance46;
			ultraGridColumn53.Header.Caption = "...";
			ultraGridColumn53.Header.VisiblePosition = 7;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn53.Width = 8;
			appearance47.Image = ((object)(resources.GetObject("appearance47.Image")));
			ultraGridColumn54.CellAppearance = appearance47;
			ultraGridColumn54.Header.Caption = "...";
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn54.Width = 8;
			appearance48.Image = ((object)(resources.GetObject("appearance48.Image")));
			ultraGridColumn55.CellAppearance = appearance48;
			ultraGridColumn55.Header.Caption = "...";
			ultraGridColumn55.Header.VisiblePosition = 10;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn55.Width = 8;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance49.BackColor = System.Drawing.Color.Transparent;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.CardAreaAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 10F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.HeaderAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.RowSelectorAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.Override.SelectedRowAppearance = appearance52;
			this.ugdGarantiaLineaCreditoAvaluoDoc.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdGarantiaLineaCreditoAvaluoDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdGarantiaLineaCreditoAvaluoDoc.Location = new System.Drawing.Point(632, 144);
			this.ugdGarantiaLineaCreditoAvaluoDoc.Name = "ugdGarantiaLineaCreditoAvaluoDoc";
			this.ugdGarantiaLineaCreditoAvaluoDoc.Size = new System.Drawing.Size(0, 144);
			this.ugdGarantiaLineaCreditoAvaluoDoc.TabIndex = 1037;
			this.ugdGarantiaLineaCreditoAvaluoDoc.Visible = false;
			this.ugdGarantiaLineaCreditoAvaluoDoc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// udsGarantiaLineaCreditoAvaluoDoc
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn33.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn34.DataType = typeof(System.UInt32);
			ultraDataColumn36.DataType = typeof(System.DateTime);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(bool);
			ultraDataColumn45.DataType = typeof(System.DateTime);
			this.udsGarantiaLineaCreditoAvaluoDoc.Band.Columns.AddRange(new object[] {
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
																																								 ultraDataColumn45});
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-1056, 232);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(2840, 8);
			this.groupBox3.TabIndex = 1044;
			this.groupBox3.TabStop = false;
			// 
			// frmLineaCreditoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(720, 542);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.ugdGarantiaLineaCreditoAvaluoDoc);
			this.Controls.Add(this.ugdLineaCreditoDestinoDoc);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.tabRequerimiento);
			this.Controls.Add(this.btnGuardar);
			this.Name = "frmLineaCreditoCrear";
			this.Text = "Nueva línea de crédito";
			this.Load += new System.EventHandler(this.frmLineaCreditoCrear_Load);
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			this.tabRequerimiento.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDestino)).EndInit();
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdBienInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoUtilizado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantiasProcentajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRenovacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAprobacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdLineaCreditoDestinoDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsLineaCreditoDestinoDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdGarantiaLineaCreditoAvaluoDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGarantiaLineaCreditoAvaluoDoc)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bAnadir = false;
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

		private void frmLineaCreditoCrear_Load(object sender, System.EventArgs e)
		{
			if(iTipo == iNuevo)
			{
				bNuevo = true;
				this.Text = "Nueva línea de crédito";
			}
			if(iTipo == iVer)
			{
				bVer = true;
				this.Text = "Ver línea de crédito";
			}
			if(iTipo == iAnadir)
			{
				bAnadir = true;
				this.Text = "Actualizar línea de crédito";
			}

			if(bNuevo)
			{
				//Coger id para identificar la linea de credito a crearse.
				idLineaCreditoTransaccion = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoIdTransaccion"));
			}
			else
			{
				idLineaCreditoTransaccion = 0;
			}
			//Cargar Banco
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoBancoCon"));
			ControlNuevo();
			//Consultar la garantia
			LineaCreditoGarantiaCon();
			//COnsultar destino
			LineaCreditoDestinoCon();
			//Cargar data Linea Credito
			LineaCreditoModel();
			//Control vista
			ControlVista();
			//Cargar combo tipo
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoTipoCon"));
			//Cargar proveedores
			this.cmbProveedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoProveedorCon"));
			//Mostrar por tipo de linea credito
			MostrarTipoNombre();
		}

		private void LineaCreditoModel()
		{
			if(idLineaCredito != 0)
			{
				//Consultar valor para ser llenados
				this.cmbBanco.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "idBanco", idLineaCredito));
				this.dtFechaAprobacion.Value = Funcion.dtEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "FechaAprobacion", idLineaCredito));
				this.dtFechaRenovacion.Value = Funcion.dtEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "FechaRenovacion", idLineaCredito));
				this.txtTasa.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "Tasa", idLineaCredito));
				this.txtGarantiasProcentajes.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "PorcentajeGarantia", idLineaCredito));
				this.txtMontoTotal.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "MontoTotal", idLineaCredito));
				this.txtNumero.Value = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "Numero", idLineaCredito));
				idLineaCreditoTransaccion = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "idLineaCreditoTransaccion", idLineaCredito));
				this.txtMontoUtilizado.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "MontoUtilizado", idLineaCredito));
				this.txtSaldo.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "Saldo", idLineaCredito));
				this.cmbTipo.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "idLineaCreditoTipo", idLineaCredito));
				this.cmbProveedor.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoModel '{0}', {1}", "idCliente", idLineaCredito));
				//Cargar garantía
				//this.ugdBienInmueble.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec LineaCreditoGarantiaCon {0}, {1}", frmLineaCreditoCrear.idLineaCreditoTransaccion, idLineaCredito));
				//Cargar destino
				//this.ugdDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec LineaCreditoDestinoCon {0}, {1}", frmLineaCreditoCrear.idLineaCreditoTransaccion, idLineaCredito));
			}
		}

		private void LineaCreditoMontoTotal()
		{
			this.txtMontoUtilizado.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoTotales {0}, {1}, {2}", frmLineaCreditoCrear.idLineaCreditoTransaccion, 1, 0));
			decimal dmontototal = 0;
			try {dmontototal = Convert.ToDecimal(this.txtMontoTotal.Value);}catch(Exception ex){ string sms = ex.Message;}
			this.txtSaldo.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoTotales {0}, {1}, {2}", frmLineaCreditoCrear.idLineaCreditoTransaccion, 2, dmontototal));
		}


		private void ControlVista()
		{
			if(bVer)
			{
				this.cmbBanco.Enabled = false;
				this.dtFechaAprobacion.Enabled = false;
				this.dtFechaRenovacion.Enabled = false;
				this.txtTasa.Enabled = false;
				this.txtGarantiasProcentajes.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.txtMontoTotal.Enabled = false;
				this.cmbTipo.Enabled = false;
				this.cmbProveedor.Enabled = false;
			}

			if(bAnadir)
			{
				this.cmbBanco.Enabled = false;
				this.dtFechaAprobacion.Enabled = false;
				this.dtFechaRenovacion.Enabled = false;
				this.txtTasa.Enabled = false;
				this.txtGarantiasProcentajes.Enabled = false;
				this.txtMontoTotal.Enabled = false;
				this.cmbTipo.Enabled = false;
				this.cmbProveedor.Enabled = false;
			}

		}

		private void LineaCreditoGarantiaCon()
		{
			this.ugdBienInmueble.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaCon {0}, {1}", 
				frmLineaCreditoCrear.idLineaCreditoTransaccion, idLineaCredito));
		}

		private void LineaCreditoDestinoCon()
		{
			this.ugdDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoCon {0} , {1}", 
				frmLineaCreditoCrear.idLineaCreditoTransaccion, idLineaCredito));
		}


		private void LineaCreditoGarantiaAvaluoDocCon()
		{
			this.ugdGarantiaLineaCreditoAvaluoDoc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoDocCargaCon {0}", 
				frmLineaCreditoCrear.idLineaCreditoTransaccion));
		}

		private void LineaCreditoDestinoDocCargaCon()
		{
			this.ugdLineaCreditoDestinoDoc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoDocCargaCon {0}", 
				frmLineaCreditoCrear.idLineaCreditoTransaccion));
		}

		private void ControlNuevo()
		{
			if(bNuevo)
			{
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}



		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida campos  
				if (!Validacion.vbComboVacio(this.cmbTipo, "Ingrese un tipo")) return;
				if((int)this.cmbTipo.Value == frmLineaCredito.idLineaCreditoTipoBanco)
				{
					if (!Validacion.vbComboVacio(this.cmbBanco, "Ingrese un Banco")) return;
				}
				if((int)this.cmbTipo.Value == frmLineaCredito.idLineaCreditoTipoProveedor)
				{
					if (!Validacion.vbComboVacio(this.cmbProveedor, "Ingrese un Proveedor")) return;
				}
				if (!Validacion.vbValidaFechasDesdeHasta(this.dtFechaAprobacion, this.dtFechaRenovacion)) return;
			  decimal dtasa = 0;
			  decimal dporcentajegarantia = 0;
				try { dtasa = Convert.ToDecimal(this.txtTasa.Value);} catch(Exception ex){string sms = ex.Message;}
				try { dporcentajegarantia = Convert.ToDecimal(this.txtGarantiasProcentajes.Value);} catch(Exception ex){string sms = ex.Message;}
				if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF,frmLineaCredito.iConfiguracionTasa, dtasa, 1)) return;
				if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF,frmLineaCredito.iConfiguracionGarPor, dporcentajegarantia, 1)) return;
			#endregion Valida campos  

			#region Cargar ultragrid
				LineaCreditoDestinoDocCargaCon();
				LineaCreditoGarantiaAvaluoDocCon();
			#endregion Cargar ultragrid

			#region Valida Garantias
			//Debería ingresar al menos un registro
			DataSet dsConfiguracion = new DataSet();
			DataView dvConfiguracion = new DataView();

			dvConfiguracion = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaCount {0}", idLineaCreditoTransaccion));
			dsConfiguracion = dvConfiguracion.Table.DataSet;
			bool bResultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
			string sMensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();

			if(bResultado == false)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida Garantias

			#region Valida destino
			//Debería ingresar al menos un registro
			dsConfiguracion = new DataSet();
			dvConfiguracion = new DataView();

			dvConfiguracion = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoCount {0}", idLineaCreditoTransaccion));
			dsConfiguracion = dvConfiguracion.Table.DataSet;
			bResultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
			sMensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();

			if(bResultado == false)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida destino

			#region	Generar rutas del documento

			string sCarpeta = "";
			string sUrlNube = "";
			string sRuta = "";
			string carpeta = "";
			string carpeta_3 = "";
			string dtFecha = "";
			string carpeta_1 = "";
			string carpeta_2 = "";
			string txtURL = "";
			string sextencion = "";
			int idGarantiaLineaCreditoAvaluoDoc = 0;
			int idLineaCreditoDestinoDoc = 0;

			carpeta_1 = @"\\192.168.1.10\Proyectos\LineaCredito";
			carpeta_2 = @"\\192.168.1.10\Proyectos\LineaCredito\"+MenuLatinium.stDirFacturacion;
				
			  using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
					{

					#region Configuracion conexion
						oConexion.Open();
						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;
					#endregion Configuracion conexion
						
					string squery= "";
					string squery2 = "";
					string squery3 = "";
					string squery4 = "";
					string squery5 = "";
					string squery6 = "";

					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						#region Numeracion
						if (bNuevo)
						{
							try
							{
								//Fijarse bien en el id de la numeración en producción porque puede ser otra.
								squery = string.Format("Exec NumeracionLocales 90, 0, 0");
								oCmdActualiza.CommandText = squery;
								this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
							}
							catch(Exception ex)
							{
								oTransaction.Rollback();
								//Funcion.WriteLog(ex.Message);
								Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
								MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								return;
							}
						}
						#endregion Numeracion

						#region Garbar linea credito

						try
						{
							if(bNuevo)
							{
								idLineaCredito = 0;
							}
							//Guardar requerimiento para luego darle seguimiento
							//int idLineaCredito = 0;
							squery2 = string.Format("Exec LineaCreditoCrear {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', {7}, {8}, {9}, {10}", 
								(int)this.cmbBanco.Value, 
								Convert.ToDateTime(this.dtFechaAprobacion.Value).ToString("yyyyMMdd"),
								Convert.ToDateTime(this.dtFechaRenovacion.Value).ToString("yyyyMMdd"),
								dtasa, 
								dporcentajegarantia,
								idLineaCreditoTransaccion,
								this.txtNumero.Text,
								idLineaCredito,
								Convert.ToDecimal(this.txtMontoTotal.Value),
								(int)this.cmbTipo.Value,
								(int)this.cmbProveedor.Value
								);

							oCmdActualiza.CommandText = squery2;
							idLineaCredito = (int)oCmdActualiza.ExecuteScalar();
						}
						catch(Exception ex)
						{
							oTransaction.Rollback();
							//Funcion.WriteLog(ex.Message);
							Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery2, Funcion.slogtipoerror);
							MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}

						
						#endregion Grabar linea credito

						#region Cargar documentos avaluo

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugdGarantiaLineaCreditoAvaluoDoc.Rows.All)	
							if ((int)dr.Cells["idGarantiaLineaCreditoAvaluoDoc"].Value != 0)
							{
								#region Variables
								sCarpeta = "";
								sUrlNube = "";
								sRuta = "";
								carpeta = "";
								carpeta_3 = "";
								dtFecha = "";
								txtURL = "";
								sRuta = dr.Cells["txtDirectorioarchivo"].Value.ToString();		
								dtFecha = DateTime.Now.ToString("yyyyMMdd");
								idGarantiaLineaCreditoAvaluoDoc = (int)dr.Cells["idGarantiaLineaCreditoAvaluoDoc"].Value;	
								#endregion Variables
								
								#region carpetas
								carpeta = @"\\192.168.1.10\Proyectos\LineaCredito\"+MenuLatinium.stDirFacturacion+@"\" + dtFecha +  "_" + this.txtNumero.Text +"_"+idLineaCredito.ToString();
								carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + dtFecha + this.txtNumero.Text+"_"+idLineaCredito.ToString();
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
								string sNombreDocuento = "";
								try
								{
									squery3 = string.Format("Exec NumeracionLocales 91, 0, 0");
									oCmdActualiza.CommandText = squery3;
									sNombreDocuento =oCmdActualiza.ExecuteScalar().ToString();
								}
								catch(Exception ex)
								{
									oTransaction.Rollback();
									//Funcion.WriteLog(ex.Message);
									Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery3, Funcion.slogtipoerror);
									MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									return;
								}
								carpeta  = (carpeta + @"\" + sNombreDocuento.Replace(" ", "_") + sextencion); // txtURL.Replace(" ", "_"));
								#region Existe Documento
								if (System.IO.File.Exists(carpeta))
								{
									MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
									return;
								}
								else
								{

									try
									{
										System.IO.File.Copy (txtDirectorioarchivo ,carpeta, false);
										sCarpeta = carpeta;
										sUrlNube = carpeta_3.Replace(" ", "_")+ @"\" + sNombreDocuento.Replace(" ", "_") + sextencion;//+ txtURL.Replace(" ", "_");
										//MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
										//Si se guardo crear un registro por el nuevo guardado
										squery4 = string.Format("Exec LineaCreditoGarantiaAvaluoDocAct {0}, '{1}', '{2}'", 
											idGarantiaLineaCreditoAvaluoDoc,
											sCarpeta,
											sUrlNube
											);
										oCmdActualiza.CommandText = squery4;
										oCmdActualiza.ExecuteNonQuery();
									}
									catch(Exception ex)
									{
										oTransaction.Rollback();
										//Funcion.WriteLog(ex.Message);
										Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery4, Funcion.slogtipoerror);
										MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										return;
									}
								}
								#endregion Existe Documento
								#endregion Generar ruta documento
							}
						#endregion Cargar documentos avaluo

						#region Cargar documentos destino
						//Recorrer la tabla para crear el file en el .10

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugdLineaCreditoDestinoDoc.Rows.All)	
							if ((int)dr.Cells["idLineaCreditoDestinoDoc"].Value != 0)
							{
								#region Variables
								sCarpeta = "";
								sUrlNube = "";
								sRuta = "";
								carpeta = "";
								carpeta_3 = "";
								dtFecha = "";
								txtURL = "";
								sRuta = dr.Cells["txtDirectorioarchivo"].Value.ToString();		
								dtFecha = DateTime.Now.ToString("yyyyMMdd");
								idLineaCreditoDestinoDoc = (int)dr.Cells["idLineaCreditoDestinoDoc"].Value;	
								#endregion Variables
								
								#region carpetas
								carpeta = @"\\192.168.1.10\Proyectos\LineaCredito\"+MenuLatinium.stDirFacturacion+@"\" + dtFecha +  "_" + this.txtNumero.Text + "_" + idLineaCredito.ToString();
								carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + dtFecha + this.txtNumero.Text + "_" + idLineaCredito.ToString();
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
								string sNombreDocuento = "";
								try
								{
									squery5 = string.Format("Exec NumeracionLocales 91, 0, 0");
									oCmdActualiza.CommandText = squery5;
									sNombreDocuento = oCmdActualiza.ExecuteScalar().ToString();
									carpeta  = (carpeta + @"\" + sNombreDocuento.Replace(" ", "_") + sextencion); // txtURL.Replace(" ", "_"));
								}
								catch(Exception ex)
								{
									oTransaction.Rollback();
									//Funcion.WriteLog(ex.Message);
									Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery5, Funcion.slogtipoerror);
									MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									return;
								}
								
								#region Existe Documento
								if (System.IO.File.Exists(carpeta))
								{
									MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
									return;
								}
								else
								{
									try
									{
										System.IO.File.Copy (txtDirectorioarchivo ,carpeta, false);
										sCarpeta = carpeta;
										squery6 = carpeta_3.Replace(" ", "_")+ @"\" + sNombreDocuento.Replace(" ", "_") + sextencion;//+ txtURL.Replace(" ", "_");
										//MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
										//Si se guardo crear un registro por el nuevo guardado
										squery6 = string.Format("Exec LineaCreditoDestinoDocAct {0}, '{1}', '{2}'", 
											idLineaCreditoDestinoDoc,
											sCarpeta,
											sUrlNube
											);
										oCmdActualiza.CommandText = squery6;
										oCmdActualiza.ExecuteNonQuery();
									}
									catch(Exception ex)
									{
										oTransaction.Rollback();
										//Funcion.WriteLog(ex.Message);
										Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery6, Funcion.slogtipoerror);
										MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										return;
									}
								}
								#endregion Existe Documento
								#endregion Generar ruta documento
							}
						#endregion Cargar documentos destino

						oTransaction.Commit();
						oConexion.Close();

						if(bNuevo)
						{
							MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se generó el siguiente número de líne de credito: {0}", this.txtNumero.Text), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						if(bAnadir)
						{
							MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se guardaron correctamente los registros.", this.txtNumero.Text), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						this.DialogResult = DialogResult.OK;

					}
					catch (Exception ex)
					{
						oTransaction.Rollback();
						this.txtNumero.Text = "";

						if(bNuevo)
						{
							idLineaCredito = 0;
						}
						//MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery + " || " + squery2 + " || "+ squery3+ " || "+ squery4+ " || "+ squery5+ " || "+ squery6, Funcion.slogtipoerror);
						MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					finally
					{
						oConexion.Close();
					}
			}
			#endregion Generar rutas del documento

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ugdDocumento_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			//e.Row.Cells["idLineaCredito"].Value = 0;
			//e.Row.Cells["UrlCompartido"].Value = "";
			//e.Row.Cells["idSeguimiento"].Value = 0;
			//e.Row.Cells["FechaDeposito"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); //DateTime.Today;
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


		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void label9_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ugdMenu_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if(bVer)
			{
				e.Layout.Bands[0].Columns["btnAnadir"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				e.Layout.Bands[0].Columns["btnEliminar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void udsGarantiaLinea_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				int idGarantiaLineaCredito = 0;
				//Nombre de las columna
				string scolumna = e.Cell.Column.ToString();
				if(scolumna == "btnAnadir")
				{
					#region Añadir garantía
//					try{idGarantiaLineaCredito = (int)e.Cell.Row.Cells["idGarantiaLineaCredito"].Value;}catch(Exception ex){string sms = ex.Message;};
					using (frmLineaCreditoGarantiaAnadir frlcc = new frmLineaCreditoGarantiaAnadir(idGarantiaLineaCredito, frmLineaCredito.iNuevo))
					{		
						if (DialogResult.OK == frlcc.ShowDialog())
						{
							LineaCreditoGarantiaCon();
						}
					}
					#endregion Añadir garantía
				}
				if(scolumna == "btnVer")
				{
					#region Ver garantía
					idGarantiaLineaCredito = (int)e.Cell.Row.Cells["idGarantiaLineaCredito"].Value;
					using (frmLineaCreditoGarantiaAnadir frlcc = new frmLineaCreditoGarantiaAnadir(idGarantiaLineaCredito, frmLineaCredito.iVer))
					{		
						if (DialogResult.OK == frlcc.ShowDialog())
						{
							LineaCreditoGarantiaCon();
						}
					}
					#endregion Ver garantía
				}
				if(scolumna == "btnEliminar")
				{
					idGarantiaLineaCredito = (int)e.Cell.Row.Cells["idGarantiaLineaCredito"].Value;
					#region Eliminar garantía
					using (frmLineaCreditoGarantiaAnadir frlcc = new frmLineaCreditoGarantiaAnadir(idGarantiaLineaCredito, frmLineaCredito.iEliminar))
					{		
						if (DialogResult.OK == frlcc.ShowDialog())
						{
							LineaCreditoGarantiaCon();
						}
					}
					#endregion Eliminar garantía
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void ugdDestino_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				//Nombre de las columna
				string scolumna = e.Cell.Column.ToString();
				int idLineaCreditoDestino = 0;

				if(scolumna == "btnAnadir")
				{
					#region Añadir destino
					using (frmLineaCreditoDestinoAnadir flcda = new frmLineaCreditoDestinoAnadir(idLineaCreditoDestino, frmLineaCreditoCrear.iNuevo))
					{		
						if (DialogResult.OK == flcda.ShowDialog())
						{
							LineaCreditoDestinoCon();
							LineaCreditoMontoTotal();
						}
					}
					#endregion Añadir destino
				}
				if(scolumna == "btnVer")
				{
					#region Ver destino
					idLineaCreditoDestino = (int)e.Cell.Row.Cells["idLineaCreditoDestino"].Value;
					using (frmLineaCreditoDestinoAnadir flcda = new frmLineaCreditoDestinoAnadir(idLineaCreditoDestino, frmLineaCreditoCrear.iVer))
					{		
						if (DialogResult.OK == flcda.ShowDialog())
						{
							LineaCreditoDestinoCon();
						}
					}
					#endregion Ver destino
				}

				if(scolumna == "btnEliminar")
				{
					idLineaCreditoDestino = (int)e.Cell.Row.Cells["idLineaCreditoDestino"].Value;
					#region Eliminar destino
					using (frmLineaCreditoDestinoAnadir flcda = new frmLineaCreditoDestinoAnadir(idLineaCreditoDestino, frmLineaCredito.iEliminar))
					{		
						if (DialogResult.OK == flcda.ShowDialog())
						{
							LineaCreditoDestinoCon();
							LineaCreditoMontoTotal();
						}
					}
					#endregion Eliminar destino
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void ugdBienInmueble_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if(bVer || bAnadir)
			{
				e.Layout.Bands[0].Columns["btnAnadir"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				e.Layout.Bands[0].Columns["btnEliminar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtMontoTotal_ValueChanged(object sender, System.EventArgs e)
		{
		LineaCreditoMontoTotal();
		}

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraCombo1_ValueChanged(object sender, System.EventArgs e)
		{
			MostrarTipoNombre();
		}

		public void MostrarTipoNombre()
		{
			this.cmbBanco.Visible = false;
			this.lblBanco.Visible = false;
			this.cmbProveedor.Visible = false;
			this.lblProveedor.Visible = false;
			int iTipo = 0;
			try {iTipo = (int)this.cmbTipo.Value;} 
			catch(Exception ex){ string sms = ex.Message;}

				//Tipo Banco
				if(iTipo == frmLineaCredito.idLineaCreditoTipoBanco)
				{
					this.cmbBanco.Visible = true;
					this.lblBanco.Visible = true;
					this.cmbProveedor.Value = 0;
				}
				//Tipo Proveedor
				if(iTipo== frmLineaCredito.idLineaCreditoTipoProveedor)
				{
					this.cmbProveedor.Visible = true;
					this.lblProveedor.Visible = true;
					this.cmbBanco.Value = 0;
				}
			

		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label12_Click(object sender, System.EventArgs e)
		{
		
		}



	}
}
