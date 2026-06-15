using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmVenta.
	/// </summary>
	public class frmVentaCrediPoint : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label61;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaDeEntrada;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcTransaccion;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtpHoraEntrega;
		private System.Windows.Forms.Label label67;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDirecciones;
		private System.Windows.Forms.Label label66;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private System.Windows.Forms.Label label72;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdSerial;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private System.Windows.Forms.Label label59;
		private System.Windows.Forms.Label label58;
		private System.Windows.Forms.Label label51;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.ComponentModel.IContainer components;
	
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidFPNC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;

		#region Variables
		bool bAsientoAutomatico = false;
		bool bImpreso = false;
		int idTipoPrecio = 0;
		bool BUsoNotaDeCredito = false;
		decimal dSaldoNotaDeCredito = 0;
		bool BRetencion = false;
		bool ModSerial = false;
		bool bEliminaAlValidar = false;
		bool bNuevo = false;
		bool bEdicion = false;
		decimal dAnticipo = 0;
		bool bUsoAnticipo = false;
		bool bPromocionCombo = false;
		bool bPromocion = false;
		bool bCombo = false;
		int IdSigno = 2;
		int idBloqueaTransacciones = 0;
		bool bSerialesL = false;
		bool bFacturacManual = false;
		bool bSerialNoExiste = false;
		bool bConsignacion = false;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1DataRow drSeteo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private System.Windows.Forms.Label lblEstado;
		bool bIncluyeIva = false;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnRetener;
		private System.Windows.Forms.Button btnCancelarRetencion;
		private Acceso miAcceso;
		private bool bActualiza = true;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieRetencion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroRetencion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizacionRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		bool bRastreo = false;
		bool bRetencion = false;
		decimal dValorRetencion = 0; 
		bool bAnticipo = false;
		decimal dValorAnticipo = 0;
		bool bCrediPoint = false;
		decimal dValorCrediPoint = 0;
		private System.Windows.Forms.Label lblUsuario;
		decimal dEntrada = 0;
		bool bNumeracionAutomatica = false;
		decimal dTotalCuotas = 0;
		int NumCuotas = 0;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbCaducidad;
		DateTime dtFechPrimPago = DateTime.Today;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.ContextMenu mnuBuscar;
		private System.Windows.Forms.MenuItem mnuBuscarPedido;
		private System.Windows.Forms.MenuItem mnuBuscarReserva;
		private System.Windows.Forms.MenuItem mnuBuscarConsignacion;
		private System.Windows.Forms.Button btnCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdproyectoZoom;
		private System.Windows.Forms.Button btnTemporal;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.TextBox txtSerial;
		private System.Windows.Forms.GroupBox gbBotones;
		bool bAnulado = false;
		int idTipoOrigen = 0;
		decimal dCuotaAsignada = 0;
		decimal dInteres = 0;
		decimal dCuotaDisponible = 0;
		decimal dCuotaUsada = 0;
		bool AutorizaEntrada = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdEF;
		private System.Windows.Forms.Button btnFE;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidRetencion;
		private System.Windows.Forms.ImageList imageList1;
		decimal dValorCuota = 0;
		#endregion Variables
		private System.Windows.Forms.Button btnImprimeDctos;
		private System.Windows.Forms.Button btnClaveDescuento;

		int idCompraV = 0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidBanco;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPlan;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPlazo;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBaseImponible;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRetencion;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblPlan;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsientoRetencion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroCredito;
		bool bDescuentoActivo = false;

		public frmVentaCrediPoint(int IdCompraV)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idCompraV = IdCompraV;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmVentaCrediPoint()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 8, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmVentaCrediPoint));
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDreccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDreccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Manual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Manual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SerieL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label72 = new System.Windows.Forms.Label();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.grdSerial = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtSerial = new System.Windows.Forms.TextBox();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label7 = new System.Windows.Forms.Label();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtAutorizacionRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumeroRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnRetener = new System.Windows.Forms.Button();
			this.btnCancelarRetencion = new System.Windows.Forms.Button();
			this.label59 = new System.Windows.Forms.Label();
			this.label58 = new System.Windows.Forms.Label();
			this.label51 = new System.Windows.Forms.Label();
			this.dtFechaRetencion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.dtpHoraEntrega = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label67 = new System.Windows.Forms.Label();
			this.cmbDirecciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label66 = new System.Windows.Forms.Label();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label61 = new System.Windows.Forms.Label();
			this.txtCuotaDeEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.utcTransaccion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidFPNC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.mnuBuscar = new System.Windows.Forms.ContextMenu();
			this.mnuBuscarPedido = new System.Windows.Forms.MenuItem();
			this.mnuBuscarReserva = new System.Windows.Forms.MenuItem();
			this.mnuBuscarConsignacion = new System.Windows.Forms.MenuItem();
			this.btnCliente = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNumeroCredito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumIdproyectoZoom = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnTemporal = new System.Windows.Forms.Button();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.txtIdEF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnFE = new System.Windows.Forms.Button();
			this.txtidRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimeDctos = new System.Windows.Forms.Button();
			this.btnClaveDescuento = new System.Windows.Forms.Button();
			this.txtidBanco = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPlan = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPlazo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtBaseImponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPlan = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdAsientoRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacionRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetencion)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).BeginInit();
			this.utcTransaccion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidFPNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPlan)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoRetencion)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.label72);
			this.ultraTabPageControl9.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl9.Controls.Add(this.grdSerial);
			this.ultraTabPageControl9.Controls.Add(this.txtSerial);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(726, 147);
			// 
			// label72
			// 
			this.label72.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 6);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(89, 16);
			this.label72.TabIndex = 153;
			this.label72.Text = "Ingrese la Serial:";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraButton1.ForeColor = System.Drawing.Color.Black;
			this.ultraButton1.Location = new System.Drawing.Point(760, 18);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 25);
			this.ultraButton1.TabIndex = 152;
			this.ultraButton1.Text = "&N";
			this.ultraButton1.Visible = false;
			// 
			// grdSerial
			// 
			this.grdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.grdSerial.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSerial.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdSerial.DisplayLayout.Appearance = appearance1;
			this.grdSerial.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 28;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 41;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 85;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 250;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 441;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 34;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 55;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
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
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.grdSerial.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSerial.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.grdSerial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdSerial.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.grdSerial.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSerial.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdSerial.Enabled = false;
			this.grdSerial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerial.Location = new System.Drawing.Point(8, 26);
			this.grdSerial.Name = "grdSerial";
			this.grdSerial.Size = new System.Drawing.Size(712, 118);
			this.grdSerial.TabIndex = 17;
			this.grdSerial.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdSerial_InitializeRow);
			this.grdSerial.DoubleClick += new System.EventHandler(this.grdSerial_DoubleClick);
			this.grdSerial.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdSerial_BeforeRowsDeleted);
			this.grdSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdSerial_KeyDown);
			this.grdSerial.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdSerial_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.Location = new System.Drawing.Point(104, 3);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(304, 20);
			this.txtSerial.TabIndex = 16;
			this.txtSerial.Text = "";
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			this.txtSerial.TextChanged += new System.EventHandler(this.txtSerial_TextChanged);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.label7);
			this.ultraTabPageControl10.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl10.Controls.Add(this.txtAutorizacionRetencion);
			this.ultraTabPageControl10.Controls.Add(this.txtNumeroRetencion);
			this.ultraTabPageControl10.Controls.Add(this.txtSerieRetencion);
			this.ultraTabPageControl10.Controls.Add(this.btnRetener);
			this.ultraTabPageControl10.Controls.Add(this.btnCancelarRetencion);
			this.ultraTabPageControl10.Controls.Add(this.label59);
			this.ultraTabPageControl10.Controls.Add(this.label58);
			this.ultraTabPageControl10.Controls.Add(this.label51);
			this.ultraTabPageControl10.Controls.Add(this.dtFechaRetencion);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(726, 147);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Enabled = false;
			this.label7.Location = new System.Drawing.Point(8, 88);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 16);
			this.label7.TabIndex = 276;
			this.label7.Text = "Fecha";
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource3;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance8;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 39;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 22;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 22;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 44;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "%";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 40;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 51;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 140;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance11;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 70;
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
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 10F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultraGrid2.Enabled = false;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(328, 0);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(392, 146);
			this.ultraGrid2.TabIndex = 259;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn12.DefaultValue = false;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			// txtAutorizacionRetencion
			// 
			this.txtAutorizacionRetencion.AcceptsTab = true;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutorizacionRetencion.Appearance = appearance19;
			this.txtAutorizacionRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizacionRetencion.Enabled = false;
			this.txtAutorizacionRetencion.Location = new System.Drawing.Point(80, 8);
			this.txtAutorizacionRetencion.MaxLength = 37;
			this.txtAutorizacionRetencion.Name = "txtAutorizacionRetencion";
			this.txtAutorizacionRetencion.Size = new System.Drawing.Size(232, 22);
			this.txtAutorizacionRetencion.TabIndex = 18;
			this.txtAutorizacionRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutorizacionRetencion_KeyPress);
			this.txtAutorizacionRetencion.ValueChanged += new System.EventHandler(this.txtAutorizacionRetencion_ValueChanged);
			// 
			// txtNumeroRetencion
			// 
			this.txtNumeroRetencion.AcceptsTab = true;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroRetencion.Appearance = appearance20;
			this.txtNumeroRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroRetencion.Enabled = false;
			this.txtNumeroRetencion.Location = new System.Drawing.Point(80, 56);
			this.txtNumeroRetencion.MaxLength = 9;
			this.txtNumeroRetencion.Name = "txtNumeroRetencion";
			this.txtNumeroRetencion.Size = new System.Drawing.Size(120, 22);
			this.txtNumeroRetencion.TabIndex = 20;
			this.txtNumeroRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumeroRetencion_KeyPress);
			this.txtNumeroRetencion.Validated += new System.EventHandler(this.txtNumeroRetencion_Validated);
			this.txtNumeroRetencion.ValueChanged += new System.EventHandler(this.txtNumeroRetencion_ValueChanged);
			// 
			// txtSerieRetencion
			// 
			this.txtSerieRetencion.AcceptsTab = true;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieRetencion.Appearance = appearance21;
			this.txtSerieRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieRetencion.Enabled = false;
			this.txtSerieRetencion.Location = new System.Drawing.Point(80, 32);
			this.txtSerieRetencion.MaxLength = 6;
			this.txtSerieRetencion.Name = "txtSerieRetencion";
			this.txtSerieRetencion.Size = new System.Drawing.Size(80, 22);
			this.txtSerieRetencion.TabIndex = 19;
			this.txtSerieRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieRetencion_KeyPress);
			this.txtSerieRetencion.ValueChanged += new System.EventHandler(this.txtSerieRetencion_ValueChanged);
			// 
			// btnRetener
			// 
			this.btnRetener.Enabled = false;
			this.btnRetener.Location = new System.Drawing.Point(8, 120);
			this.btnRetener.Name = "btnRetener";
			this.btnRetener.Size = new System.Drawing.Size(104, 23);
			this.btnRetener.TabIndex = 21;
			this.btnRetener.Text = "&Crear Retención";
			this.btnRetener.Click += new System.EventHandler(this.btnRetener_Click);
			// 
			// btnCancelarRetencion
			// 
			this.btnCancelarRetencion.Enabled = false;
			this.btnCancelarRetencion.Location = new System.Drawing.Point(128, 120);
			this.btnCancelarRetencion.Name = "btnCancelarRetencion";
			this.btnCancelarRetencion.Size = new System.Drawing.Size(104, 23);
			this.btnCancelarRetencion.TabIndex = 22;
			this.btnCancelarRetencion.Text = "&Borrar Retención";
			this.btnCancelarRetencion.Click += new System.EventHandler(this.btnCancelarRetencion_Click);
			// 
			// label59
			// 
			this.label59.AutoSize = true;
			this.label59.BackColor = System.Drawing.Color.Transparent;
			this.label59.Enabled = false;
			this.label59.Location = new System.Drawing.Point(8, 59);
			this.label59.Name = "label59";
			this.label59.Size = new System.Drawing.Size(44, 16);
			this.label59.TabIndex = 139;
			this.label59.Text = "Número";
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.BackColor = System.Drawing.Color.Transparent;
			this.label58.Enabled = false;
			this.label58.Location = new System.Drawing.Point(8, 11);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(65, 16);
			this.label58.TabIndex = 138;
			this.label58.Text = "Autorización";
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Enabled = false;
			this.label51.Location = new System.Drawing.Point(8, 35);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(29, 16);
			this.label51.TabIndex = 136;
			this.label51.Text = "Serie";
			// 
			// dtFechaRetencion
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRetencion.Appearance = appearance22;
			this.dtFechaRetencion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaRetencion.DateButtons.Add(dateButton1);
			this.dtFechaRetencion.Enabled = false;
			this.dtFechaRetencion.Format = "dd/MM/yyyy";
			this.dtFechaRetencion.Location = new System.Drawing.Point(80, 88);
			this.dtFechaRetencion.Name = "dtFechaRetencion";
			this.dtFechaRetencion.NonAutoSizeHeight = 23;
			this.dtFechaRetencion.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRetencion.SpinButtonsVisible = true;
			this.dtFechaRetencion.TabIndex = 275;
			this.dtFechaRetencion.Value = ((object)(resources.GetObject("dtFechaRetencion.Value")));
			this.dtFechaRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaRetencion_KeyPress);
			this.dtFechaRetencion.Validated += new System.EventHandler(this.dtFechaRetencion_Validated);
			this.dtFechaRetencion.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaRetencion_BeforeDropDown);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.cmbCaducidad);
			this.ultraTabPageControl2.Controls.Add(this.label2);
			this.ultraTabPageControl2.Controls.Add(this.txtSerie);
			this.ultraTabPageControl2.Controls.Add(this.label33);
			this.ultraTabPageControl2.Controls.Add(this.txtAutFactura);
			this.ultraTabPageControl2.Controls.Add(this.label31);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(726, 147);
			// 
			// cmbCaducidad
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCaducidad.Appearance = appearance23;
			this.cmbCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbCaducidad.DateButtons.Add(dateButton2);
			this.cmbCaducidad.Enabled = false;
			this.cmbCaducidad.Format = "dd/MM/yyyy";
			this.cmbCaducidad.Location = new System.Drawing.Point(112, 80);
			this.cmbCaducidad.Name = "cmbCaducidad";
			this.cmbCaducidad.NonAutoSizeHeight = 23;
			this.cmbCaducidad.Size = new System.Drawing.Size(112, 21);
			this.cmbCaducidad.SpinButtonsVisible = true;
			this.cmbCaducidad.TabIndex = 25;
			this.cmbCaducidad.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.cmbCaducidad.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.cmbCaducidad_BeforeDropDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(8, 83);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 16);
			this.label2.TabIndex = 100;
			this.label2.Text = "Caduca";
			// 
			// txtSerie
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance24;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Enabled = false;
			this.txtSerie.Location = new System.Drawing.Point(112, 16);
			this.txtSerie.MaxLength = 10;
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(96, 22);
			this.txtSerie.TabIndex = 23;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Enabled = false;
			this.label33.Location = new System.Drawing.Point(8, 19);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(29, 16);
			this.label33.TabIndex = 97;
			this.label33.Text = "Serie";
			// 
			// txtAutFactura
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance25;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(112, 48);
			this.txtAutFactura.MaxLength = 10;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(208, 22);
			this.txtAutFactura.TabIndex = 24;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Enabled = false;
			this.label31.Location = new System.Drawing.Point(8, 51);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 95;
			this.label31.Text = "Autorización";
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.dtpHoraEntrega);
			this.pgEntrega.Controls.Add(this.label67);
			this.pgEntrega.Controls.Add(this.cmbDirecciones);
			this.pgEntrega.Controls.Add(this.label66);
			this.pgEntrega.Controls.Add(this.chkEntregado);
			this.pgEntrega.Controls.Add(this.dtFechaEntrega);
			this.pgEntrega.Controls.Add(this.label19);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(726, 147);
			// 
			// dtpHoraEntrega
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHoraEntrega.Appearance = appearance26;
			this.dtpHoraEntrega.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtpHoraEntrega.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never;
			this.dtpHoraEntrega.Enabled = false;
			this.dtpHoraEntrega.FormatString = "HH:mm";
			this.dtpHoraEntrega.Location = new System.Drawing.Point(88, 72);
			this.dtpHoraEntrega.MaskInput = "{LOC}hh:mm";
			this.dtpHoraEntrega.Name = "dtpHoraEntrega";
			this.dtpHoraEntrega.Size = new System.Drawing.Size(80, 22);
			this.dtpHoraEntrega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtpHoraEntrega.TabIndex = 28;
			this.dtpHoraEntrega.Value = null;
			this.dtpHoraEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpHoraEntrega_KeyPress);
			this.dtpHoraEntrega.ValueChanged += new System.EventHandler(this.dtpHoraEntrega_ValueChanged);
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Enabled = false;
			this.label67.Location = new System.Drawing.Point(16, 72);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 16);
			this.label67.TabIndex = 128;
			this.label67.Text = "Hora";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDirecciones
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDirecciones.Appearance = appearance27;
			this.cmbDirecciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDirecciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDirecciones.DataSource = this.ultraDataSource1;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 360;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbDirecciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbDirecciones.DisplayMember = "Direccion";
			this.cmbDirecciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDirecciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDirecciones.Enabled = false;
			this.cmbDirecciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDirecciones.Location = new System.Drawing.Point(88, 104);
			this.cmbDirecciones.Name = "cmbDirecciones";
			this.cmbDirecciones.Size = new System.Drawing.Size(536, 21);
			this.cmbDirecciones.TabIndex = 29;
			this.cmbDirecciones.ValueMember = "idDreccion";
			this.cmbDirecciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbDirecciones_KeyPress);
			this.cmbDirecciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDirecciones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Enabled = false;
			this.label66.Location = new System.Drawing.Point(16, 104);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(50, 16);
			this.label66.TabIndex = 126;
			this.label66.Text = "Dirección";
			// 
			// chkEntregado
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance28;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.Enabled = false;
			this.chkEntregado.Location = new System.Drawing.Point(16, 8);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(120, 22);
			this.chkEntregado.TabIndex = 26;
			this.chkEntregado.Text = "Entrega a domicilio";
			this.chkEntregado.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkEntregado_KeyPress);
			this.chkEntregado.CheckedChanged += new System.EventHandler(this.chkEntregado_CheckedChanged);
			// 
			// dtFechaEntrega
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance29;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton3);
			this.dtFechaEntrega.Enabled = false;
			this.dtFechaEntrega.Format = "dd/MM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(88, 40);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 27;
			this.dtFechaEntrega.Value = new System.DateTime(2005, 7, 1, 0, 0, 0, 0);
			this.dtFechaEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaEntrega_KeyPress);
			this.dtFechaEntrega.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaEntrega_BeforeDropDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Enabled = false;
			this.label19.Location = new System.Drawing.Point(16, 40);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(34, 16);
			this.label19.TabIndex = 44;
			this.label19.Text = "Fecha";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(726, 147);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.AcceptsTab = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance30;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(7, 9);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(713, 135);
			this.txtNotas.TabIndex = 30;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(264, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 128;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(488, 11);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 151;
			this.lblVendedor.Text = "Vendedor";
			// 
			// cmbVendedor
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance31;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 335;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(552, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(256, 21);
			this.cmbVendedor.TabIndex = 4;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 43);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 16);
			this.lblIdentificacion.TabIndex = 161;
			this.lblIdentificacion.Text = "Identificación";
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(264, 43);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 16);
			this.lblCliente.TabIndex = 160;
			this.lblCliente.Text = "N. Identificación";
			// 
			// cmbTipoRuc
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance32;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 233;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 40);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 6;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoRuc_KeyPress);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance33;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(360, 40);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 7;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance34;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(552, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 8;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(264, 75);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(43, 16);
			this.label61.TabIndex = 164;
			this.label61.Text = "Entrada";
			// 
			// txtCuotaDeEntrada
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDeEntrada.Appearance = appearance35;
			this.txtCuotaDeEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDeEntrada.Enabled = false;
			this.txtCuotaDeEntrada.FormatString = "#,##0.00";
			this.txtCuotaDeEntrada.Location = new System.Drawing.Point(360, 72);
			this.txtCuotaDeEntrada.Name = "txtCuotaDeEntrada";
			this.txtCuotaDeEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDeEntrada.PromptChar = ' ';
			this.txtCuotaDeEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtCuotaDeEntrada.TabIndex = 12;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 75);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 167;
			this.label6.Text = "Fecha";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance36.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance36;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 50;
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 120;
			ultraGridColumn30.Header.VisiblePosition = 4;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.Header.VisiblePosition = 5;
			ultraGridColumn31.Width = 239;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.Header.VisiblePosition = 6;
			ultraGridColumn32.Width = 268;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance37;
			ultraGridColumn33.Format = "";
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.PromptChar = ' ';
			ultraGridColumn33.Width = 85;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance38;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn34.Width = 96;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance39;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "IVA";
			ultraGridColumn35.Header.VisiblePosition = 9;
			ultraGridColumn35.PromptChar = ' ';
			ultraGridColumn35.Width = 64;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance40;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 10;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.PromptChar = ' ';
			ultraGridColumn36.Width = 90;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance41;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "% Desc.";
			ultraGridColumn37.Header.VisiblePosition = 11;
			ultraGridColumn37.PromptChar = ' ';
			ultraGridColumn37.Width = 75;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance42;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Subtotal";
			ultraGridColumn38.Header.VisiblePosition = 12;
			ultraGridColumn38.PromptChar = ' ';
			ultraGridColumn38.Width = 96;
			ultraGridColumn39.Header.VisiblePosition = 13;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 14;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 15;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 16;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 17;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 18;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 19;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 20;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 63;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 21;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 22;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 23;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance43;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 24;
			ultraGridColumn50.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50});
			ultraGridBand6.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance44;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance45;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance46.FontData.BoldAsString = "True";
			appearance46.FontData.Name = "Arial";
			appearance46.FontData.SizeInPoints = 10F;
			appearance46.ForeColor = System.Drawing.Color.White;
			appearance46.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance46;
			appearance47.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance47.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(9, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1007, 208);
			this.ultraGrid1.TabIndex = 14;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(bool);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn28.DefaultValue = 0;
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn29.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn30.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn33.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(bool);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(bool);
			ultraDataColumn42.DefaultValue = false;
			ultraDataColumn43.DataType = typeof(bool);
			ultraDataColumn43.DefaultValue = false;
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn32,
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
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(832, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 190;
			this.lblNumero.Text = "Número";
			// 
			// txtNumero
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance50;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(912, 7);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 5;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(752, 344);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(46, 16);
			this.lblSubTotal.TabIndex = 208;
			this.lblSubTotal.Text = "Subtotal";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(752, 419);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 206;
			this.lblIva.Text = "IVA";
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(752, 368);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(58, 16);
			this.lblDescuento.TabIndex = 204;
			this.lblDescuento.Text = "Descuento";
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva0.Location = new System.Drawing.Point(888, 320);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 200;
			this.lblIva0.Text = "IVA 0";
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(992, 320);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 16);
			this.label55.TabIndex = 203;
			this.label55.Text = "IVA";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance51;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(840, 440);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 34);
			this.txtTotal.TabIndex = 222;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance52;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(936, 416);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 221;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance53;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(840, 368);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 219;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance54;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(936, 368);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 216;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance55;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(936, 344);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 214;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance56;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(840, 344);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 212;
			// 
			// txtNumIdAsiento
			// 
			this.txtNumIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdAsiento.Enabled = false;
			this.txtNumIdAsiento.Location = new System.Drawing.Point(952, 496);
			this.txtNumIdAsiento.Name = "txtNumIdAsiento";
			this.txtNumIdAsiento.PromptChar = ' ';
			this.txtNumIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdAsiento.TabIndex = 210;
			this.txtNumIdAsiento.Visible = false;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(752, 448);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 223;
			this.label1.Text = "Total";
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance57;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.pgEntrega);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl9);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl10);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl2);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 320);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(728, 168);
			this.utcTransaccion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcTransaccion.TabIndex = 15;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Seriales";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Retenciones";
			ultraTab3.TabPage = this.ultraTabPageControl2;
			ultraTab3.Text = "S. R. I.";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Entrega";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Notas";
			this.utcTransaccion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3,
																																														 ultraTab4,
																																														 ultraTab5});
			this.utcTransaccion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcTransaccion.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcTransaccion_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(726, 147);
			// 
			// cmbFormaPago
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance58;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 152;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn51,
																										 ultraGridColumn52});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 8);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 21);
			this.cmbFormaPago.TabIndex = 2;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 11);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 225;
			this.label8.Text = "Forma de Pago";
			// 
			// txtComprobante
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance59;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(912, 40);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(104, 22);
			this.txtComprobante.TabIndex = 10;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(832, 43);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 226;
			this.lblComprobante.Text = "Comprobante";
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
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(936, 496);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 228;
			this.txtNumIdCliente.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(904, 496);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 229;
			this.txtNumIdCompra.Visible = false;
			// 
			// txtNumidOrigen
			// 
			this.txtNumidOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen.Enabled = false;
			this.txtNumidOrigen.Location = new System.Drawing.Point(920, 496);
			this.txtNumidOrigen.Name = "txtNumidOrigen";
			this.txtNumidOrigen.PromptChar = ' ';
			this.txtNumidOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen.TabIndex = 230;
			this.txtNumidOrigen.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(968, 496);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 231;
			this.txtNumContadoCredito.Visible = false;
			// 
			// txtNumidOrigen1
			// 
			this.txtNumidOrigen1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen1.Enabled = false;
			this.txtNumidOrigen1.Location = new System.Drawing.Point(984, 496);
			this.txtNumidOrigen1.Name = "txtNumidOrigen1";
			this.txtNumidOrigen1.PromptChar = ' ';
			this.txtNumidOrigen1.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen1.TabIndex = 232;
			this.txtNumidOrigen1.Visible = false;
			// 
			// txtNumidFPNC
			// 
			this.txtNumidFPNC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidFPNC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidFPNC.Enabled = false;
			this.txtNumidFPNC.Location = new System.Drawing.Point(1000, 496);
			this.txtNumidFPNC.Name = "txtNumidFPNC";
			this.txtNumidFPNC.PromptChar = ' ';
			this.txtNumidFPNC.Size = new System.Drawing.Size(16, 22);
			this.txtNumidFPNC.TabIndex = 233;
			this.txtNumidFPNC.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(904, 520);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 234;
			this.txtIdSucursal.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(920, 520);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 235;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(936, 520);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 236;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(952, 520);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 237;
			this.txtNumEstado.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(968, 520);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 238;
			this.txtIdComprobante.Visible = false;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(984, 520);
			this.txtNumIdProyecto.Name = "txtNumIdProyecto";
			this.txtNumIdProyecto.PromptChar = ' ';
			this.txtNumIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdProyecto.TabIndex = 239;
			this.txtNumIdProyecto.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(840, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 240;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn53.Header.VisiblePosition = 2;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 180;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(360, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 3;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(488, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 244;
			this.label3.Text = "Nombre";
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(913, 512);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 39;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(642, 512);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 37;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(563, 512);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 36;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 512);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 31;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 512);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 32;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(164, 512);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 33;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 512);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 512);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(752, 320);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 254;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Enabled = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.ImageIndex = 67;
			this.btnBuscar.ImageList = this.imageList1;
			this.btnBuscar.Location = new System.Drawing.Point(717, 512);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 38;
			this.btnBuscar.Text = "BI";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// mnuBuscar
			// 
			this.mnuBuscar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.mnuBuscarPedido,
																																							this.mnuBuscarReserva,
																																							this.mnuBuscarConsignacion});
			// 
			// mnuBuscarPedido
			// 
			this.mnuBuscarPedido.Index = 0;
			this.mnuBuscarPedido.Text = "Pedido";
			this.mnuBuscarPedido.Click += new System.EventHandler(this.mnuBuscarPedido_Click);
			// 
			// mnuBuscarReserva
			// 
			this.mnuBuscarReserva.Index = 1;
			this.mnuBuscarReserva.Text = "Reserva";
			this.mnuBuscarReserva.Click += new System.EventHandler(this.mnuBuscarReserva_Click);
			// 
			// mnuBuscarConsignacion
			// 
			this.mnuBuscarConsignacion.Index = 2;
			this.mnuBuscarConsignacion.Text = "Consignación";
			this.mnuBuscarConsignacion.Visible = false;
			// 
			// btnCliente
			// 
			this.btnCliente.CausesValidation = false;
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(800, 40);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 9;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(488, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(54, 16);
			this.label4.TabIndex = 259;
			this.label4.Text = "Credito N.";
			// 
			// txtNumeroCredito
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroCredito.Appearance = appearance60;
			this.txtNumeroCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroCredito.Enabled = false;
			this.txtNumeroCredito.Location = new System.Drawing.Point(552, 72);
			this.txtNumeroCredito.Name = "txtNumeroCredito";
			this.txtNumeroCredito.Size = new System.Drawing.Size(248, 22);
			this.txtNumeroCredito.TabIndex = 13;
			// 
			// txtNumIdproyectoZoom
			// 
			this.txtNumIdproyectoZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdproyectoZoom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdproyectoZoom.Enabled = false;
			this.txtNumIdproyectoZoom.Location = new System.Drawing.Point(1000, 520);
			this.txtNumIdproyectoZoom.Name = "txtNumIdproyectoZoom";
			this.txtNumIdproyectoZoom.PromptChar = ' ';
			this.txtNumIdproyectoZoom.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdproyectoZoom.TabIndex = 260;
			this.txtNumIdproyectoZoom.Visible = false;
			// 
			// btnTemporal
			// 
			this.btnTemporal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnTemporal.CausesValidation = false;
			this.btnTemporal.Enabled = false;
			this.btnTemporal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTemporal.ImageIndex = 66;
			this.btnTemporal.ImageList = this.imageList1;
			this.btnTemporal.Location = new System.Drawing.Point(402, 512);
			this.btnTemporal.Name = "btnTemporal";
			this.btnTemporal.TabIndex = 34;
			this.btnTemporal.Text = "Temporal";
			this.btnTemporal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTemporal.Click += new System.EventHandler(this.btnTemporal_Click);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(480, 512);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 35;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.gbBotones.Location = new System.Drawing.Point(0, 496);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1024, 8);
			this.gbBotones.TabIndex = 263;
			this.gbBotones.TabStop = false;
			// 
			// txtIdEF
			// 
			this.txtIdEF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdEF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdEF.Enabled = false;
			this.txtIdEF.Location = new System.Drawing.Point(888, 496);
			this.txtIdEF.Name = "txtIdEF";
			this.txtIdEF.PromptChar = ' ';
			this.txtIdEF.Size = new System.Drawing.Size(16, 22);
			this.txtIdEF.TabIndex = 264;
			this.txtIdEF.Visible = false;
			// 
			// btnFE
			// 
			this.btnFE.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFE.CausesValidation = false;
			this.btnFE.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFE.Location = new System.Drawing.Point(736, 312);
			this.btnFE.Name = "btnFE";
			this.btnFE.Size = new System.Drawing.Size(48, 23);
			this.btnFE.TabIndex = 265;
			this.btnFE.Text = "&FE";
			this.btnFE.Visible = false;
			this.btnFE.Click += new System.EventHandler(this.btnFE_Click);
			// 
			// txtidRetencion
			// 
			this.txtidRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidRetencion.Enabled = false;
			this.txtidRetencion.Location = new System.Drawing.Point(888, 520);
			this.txtidRetencion.Name = "txtidRetencion";
			this.txtidRetencion.PromptChar = ' ';
			this.txtidRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtidRetencion.TabIndex = 266;
			this.txtidRetencion.Visible = false;
			// 
			// btnImprimeDctos
			// 
			this.btnImprimeDctos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimeDctos.CausesValidation = false;
			this.btnImprimeDctos.Enabled = false;
			this.btnImprimeDctos.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimeDctos.Image")));
			this.btnImprimeDctos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimeDctos.Location = new System.Drawing.Point(800, 512);
			this.btnImprimeDctos.Name = "btnImprimeDctos";
			this.btnImprimeDctos.Size = new System.Drawing.Size(110, 23);
			this.btnImprimeDctos.TabIndex = 267;
			this.btnImprimeDctos.Text = "Dctos CrediPoint";
			this.btnImprimeDctos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimeDctos.Click += new System.EventHandler(this.btnImprimeDctos_Click);
			// 
			// btnClaveDescuento
			// 
			this.btnClaveDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnClaveDescuento.CausesValidation = false;
			this.btnClaveDescuento.Enabled = false;
			this.btnClaveDescuento.Location = new System.Drawing.Point(991, 512);
			this.btnClaveDescuento.Name = "btnClaveDescuento";
			this.btnClaveDescuento.Size = new System.Drawing.Size(24, 22);
			this.btnClaveDescuento.TabIndex = 268;
			this.btnClaveDescuento.Text = "D";
			this.btnClaveDescuento.Visible = false;
			this.btnClaveDescuento.Click += new System.EventHandler(this.btnClaveDescuento_Click);
			// 
			// txtidBanco
			// 
			this.txtidBanco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidBanco.Enabled = false;
			this.txtidBanco.Location = new System.Drawing.Point(744, 472);
			this.txtidBanco.Name = "txtidBanco";
			this.txtidBanco.PromptChar = ' ';
			this.txtidBanco.Size = new System.Drawing.Size(16, 22);
			this.txtidBanco.TabIndex = 272;
			this.txtidBanco.Visible = false;
			// 
			// txtidPlan
			// 
			this.txtidPlan.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidPlan.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPlan.Enabled = false;
			this.txtidPlan.Location = new System.Drawing.Point(776, 472);
			this.txtidPlan.Name = "txtidPlan";
			this.txtidPlan.PromptChar = ' ';
			this.txtidPlan.Size = new System.Drawing.Size(16, 22);
			this.txtidPlan.TabIndex = 271;
			this.txtidPlan.Visible = false;
			// 
			// txtidTarjeta
			// 
			this.txtidTarjeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTarjeta.Enabled = false;
			this.txtidTarjeta.Location = new System.Drawing.Point(760, 472);
			this.txtidTarjeta.Name = "txtidTarjeta";
			this.txtidTarjeta.PromptChar = ' ';
			this.txtidTarjeta.Size = new System.Drawing.Size(16, 22);
			this.txtidTarjeta.TabIndex = 270;
			this.txtidTarjeta.Visible = false;
			// 
			// txtidPlazo
			// 
			this.txtidPlazo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPlazo.Enabled = false;
			this.txtidPlazo.Location = new System.Drawing.Point(792, 472);
			this.txtidPlazo.Name = "txtidPlazo";
			this.txtidPlazo.PromptChar = ' ';
			this.txtidPlazo.Size = new System.Drawing.Size(16, 22);
			this.txtidPlazo.TabIndex = 269;
			this.txtidPlazo.Visible = false;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(752, 395);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(174, 16);
			this.label5.TabIndex = 273;
			this.label5.Text = "BASE IMPONIBLE RETENCION";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBaseImponible
			// 
			this.txtBaseImponible.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBaseImponible.Appearance = appearance61;
			this.txtBaseImponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBaseImponible.Enabled = false;
			this.txtBaseImponible.FormatString = "#,##0.00";
			this.txtBaseImponible.Location = new System.Drawing.Point(936, 392);
			this.txtBaseImponible.Name = "txtBaseImponible";
			this.txtBaseImponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBaseImponible.PromptChar = ' ';
			this.txtBaseImponible.Size = new System.Drawing.Size(80, 22);
			this.txtBaseImponible.TabIndex = 274;
			// 
			// lblPlan
			// 
			this.lblPlan.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPlan.AutoSize = true;
			this.lblPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPlan.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPlan.Location = new System.Drawing.Point(840, 480);
			this.lblPlan.Name = "lblPlan";
			this.lblPlan.Size = new System.Drawing.Size(0, 16);
			this.lblPlan.TabIndex = 275;
			this.lblPlan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance62;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton4);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 72);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 11;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtIdAsientoRetencion
			// 
			this.txtIdAsientoRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsientoRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsientoRetencion.Enabled = false;
			this.txtIdAsientoRetencion.Location = new System.Drawing.Point(808, 472);
			this.txtIdAsientoRetencion.Name = "txtIdAsientoRetencion";
			this.txtIdAsientoRetencion.PromptChar = ' ';
			this.txtIdAsientoRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsientoRetencion.TabIndex = 276;
			this.txtIdAsientoRetencion.Visible = false;
			// 
			// frmVentaCrediPoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1024, 542);
			this.Controls.Add(this.txtIdAsientoRetencion);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblPlan);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtBaseImponible);
			this.Controls.Add(this.txtidBanco);
			this.Controls.Add(this.txtidPlan);
			this.Controls.Add(this.txtidTarjeta);
			this.Controls.Add(this.txtidPlazo);
			this.Controls.Add(this.btnClaveDescuento);
			this.Controls.Add(this.btnImprimeDctos);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.txtidRetencion);
			this.Controls.Add(this.btnFE);
			this.Controls.Add(this.txtIdEF);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnTemporal);
			this.Controls.Add(this.txtNumIdproyectoZoom);
			this.Controls.Add(this.txtNumeroCredito);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtNumIdProyecto);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumidFPNC);
			this.Controls.Add(this.txtNumidOrigen1);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumidOrigen);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.utcTransaccion);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtNumIdAsiento);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtCuotaDeEntrada);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmbFormaPago);
			this.KeyPreview = true;
			this.Name = "frmVentaCrediPoint";
			this.Text = "Ventas De CrediPOINT";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmVenta_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmVenta_Closing);
			this.Load += new System.EventHandler(this.frmVenta_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizacionRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetencion)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).EndInit();
			this.utcTransaccion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidFPNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPlan)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsientoRetencion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void Actualiza (int idCliente)
		{
			this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", idCliente));
		}

		private void HabilitaTab(bool bEstado)
		{
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in utcTransaccion.Tabs)
			{
				foreach(Control cn in ctr.TabPage.Controls) 
				{
					cn.Enabled = bEstado;
				}
			}
		}

		private void ModificaGrilla(bool bModifica)
		{
			this.ultraGrid1.Enabled = true;
			if (bModifica)
			{
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void ModificaGrillaSerial(bool bModifica)
		{
			this.grdSerial.Enabled = true;
			if (bModifica)
			{
				this.grdSerial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.grdSerial.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.grdSerial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.grdSerial.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void frmVenta_Load(object sender, System.EventArgs e)
		{	
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			
			string stTipo = "0703";
			string stPermiso = "118";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso,cdsSeteo))
			{				
				MessageBox.Show("No puede ingresar a Ventas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				#region Incluye Iva 
				if ((bool) drSeteoF["IncluyeIva"]) bIncluyeIva = true;			
				#endregion Incluye Iva

				#region Ficha de artículos
				//			if (!miAcceso.BFichaArticulos) this.btBuscaArticulo.Visible = false;
				//			if (!miAcceso.BAuditoria) this.btAuditoria.Visible = false;
				#endregion Ficha de artículos 

				IdSigno = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 1");

				drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
				drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
				drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			
				if ((bool) drSeteoF["IncluyeIva"]) bIncluyeIva = true;
			
				bSerialesL = (bool) drSeteoF["SerieLinea"];
				bFacturacManual = (bool) drSeteoF["FacturaManual"];
				bSerialNoExiste = (bool) drSeteoF["SerialNoExiste"];

				if ((bool)drSeteoF["CotizaServicios"] && !(bool)drSeteoF["CotizaInventario"])  
				{
					this.utcTransaccion.Tabs[0].Visible = false;
				}

				#region Listas
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga 1, 0");
				string sSQL = string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario);
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
				#endregion Listas

				if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
				if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;
				if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
				if (miAcceso.CambiarDescuento) this.btnClaveDescuento.Visible = true;

				ModificaGrilla(false);
				HabilitaTab(false);
				this.grdSerial.Enabled = true;
				ModificaGrillaSerial(false);

				sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if ((int)this.txtBodegaPredef.Value > 1)
				{
					this.utcTransaccion.Tabs[2].Visible = false;
					this.utcTransaccion.Tabs[3].Visible = false;
				}

				this.txtIdTipoFactura.Value = 1;
				this.txtIdComprobante.Value = 15;
				//this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

				bAsientoAutomatico = Funcion.bEjecutaSQL(cdsSeteoF, "Exec AsientoAutomaticoVerifica 1");

				if (idCompraV > 0) this.Consultar(idCompraV);
			}
		}

		private void utcTransaccion_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (!bNumeracionAutomatica) this.txtNumero.Focus();
				else this.txtRuc.Focus();
			}
		}

		private void cmbTipoRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtRuc.Focus();
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void EstadoRetencion(bool bEstado)
		{
			this.btnRetener.Enabled = bEstado;
			this.btnCancelarRetencion.Enabled = bEstado;
			this.txtSerieRetencion.Enabled = bEstado;
			this.txtNumeroRetencion.Enabled = bEstado;
			this.txtAutorizacionRetencion.Enabled = bEstado;
			this.dtFechaRetencion.Enabled = bEstado;
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13)
			{
				if (!miAcceso.CambiarFecha)
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
				else this.dtFecha.Focus();
			}
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{			
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
				if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;
			}
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value, (int)this.cmbTipoRuc.Value, (int) this.cmbBodega.Value, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtNumIdCliente.Value = miBusqueda.idCliente;
						this.cmbTipoRuc.Value = miBusqueda.idTipoRuc;
						this.txtRuc.Text = miBusqueda.Ruc;
						this.txtNombre.Text = miBusqueda.Nombre;
												
						if ((int)this.cmbTipoRuc.Value == 4) EstadoRetencion(true);
						else EstadoRetencion(false);

						if (!miAcceso.CambiarFecha) this.ultraGrid1.Focus();
						else this.dtFecha.Focus();
						
						Actualiza((int)this.txtNumIdCliente.Value);

						#region Anticipos
						dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
						if (dAnticipo > 0)
						{
							if (DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
								bUsoAnticipo = true;
							else dAnticipo = 0;
						}
						#endregion Anticipos
					}
				}
			}
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region F3
						case (int) Keys.F3:
							#region Inventario
							string sArticulo = ""; 
							sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();							
							if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO") && bEdicion && !bFormaPagoPromocional)
							{
								#region Abre Inventario
								if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
								{
									if (!(bool)this.ultraGrid1.ActiveRow.Cells["Principal"].Value) 
									{
										#region Valida Tipo de Producto Normal - Combo - Promocion
										if ((int)this.ultraGrid1.ActiveRow.Cells["Tipo"].Value > 0)
										{
											#region Modifica Articulo de Combo o Promocion
											if (miAcceso.BModArtCmbProm)
											{
												sArticulo = ""; 
												sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
																											
												using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value, Convert.ToDateTime(this.dtFecha.Value), false))
												{
													if (DialogResult.OK == miBusqueda.ShowDialog())
													{
														ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
													}
												}		
											}
											else
											{
												MessageBox.Show("No puede modificar un Producto Que Pertenece a un Combo o Promoción \n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
											}
											#endregion Modifica Articulo de Combo o Promocion
										}
										else if ((int)this.ultraGrid1.ActiveRow.Cells["Tipo"].Value == 0)
										{
											sArticulo = ""; 
											sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();											
															
											using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value, Convert.ToDateTime(this.dtFecha.Value), false))
											{
												if (DialogResult.OK == miBusqueda.ShowDialog())
												{
													ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
												}
											}												
										}
										#endregion Valida Tipo de Producto Normal - Combo - Promocion
									}										
									else
									{
										MessageBox.Show("No Puede Modificar un Producto Principal de un Combo o Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									}										
								}
								else
								{
									sArticulo = ""; 
									sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();										
															
									using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value, Convert.ToDateTime(this.dtFecha.Value), false))
									{
										if (DialogResult.OK == miBusqueda.ShowDialog())
										{
											ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
										}
									}										
								}
								#endregion Abre Inventario
							}					
							break;
							#endregion Inventario												
							#endregion F3						
							#region Flechas
						case (int) Keys.Up:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;

							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (ultraGrid1.ActiveCell.DroppedDown) return;
			
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
							//							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Iva
			if (!(bool)drSeteoF["CambiarIva"]) e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;
			#endregion Iva
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["DescuentoArt"].Format = stFormato;

			this.txtDescIvaTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtIva0.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			this.txtDescIva0Total.FormatString = stFormato;			
			this.txtDescIvaTotal.MaskInput = stInput;			
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtIva0.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			this.txtDescIva0Total.MaskInput = stInput;
			#endregion			
			#region Decimales para el precio
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;			
			#endregion Decimales para el precio
			#region Decimales para el Impuesto
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;
			#endregion
			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPorcDesc"]);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];
			stInput = stMascara + stNumero.Substring(0, iDecPorD);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			#endregion
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			
		}

		#region Funciones del grid
		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}

		private string revisaCadena(string cadenaOrg)
		{
			string letra = "";
			string stL = "'";
			string cadena = cadenaOrg;
			string cadena1 = "";
			for (int i=0;i<cadena.Length;i++)
			{
				letra = cadena.Substring(i,1);				
				if(letra == stL)
				{					
					cadena1 += "-";
				}
				else
				{
					cadena1 += letra;
				}
			}
			//				MessageBox.Show(" "+cadena1);
			return cadena1;
		}

		private decimal SumaPrecios()
		{
			decimal dPrecio = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Tipo"].Value != System.DBNull.Value)	
				{
					if (dr.Cells["Subtotal"].Value != System.DBNull.Value)
					{
						if ((int)dr.Cells["Tipo"].Value != 3)	dPrecio = dPrecio + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
					}
				}
			}

			dPrecio = dPrecio * 1.12m;

			return dPrecio;
		}

		private void Total()
		{
			#region Creacion de Variables					
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;

			#region Calculo Forma Pago Mixto 
			decimal dSubTotal1 = 0.00m;
			decimal dSubTotal2 = 0.00m;
			decimal dSubTotal3 = 0.00m;
			decimal dSubTotal4 = 0.00m;
			#endregion Calculo Forma Pago Mixto 

			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;
			this.txtIva0.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Subtotal"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100.00m, 2); //Math.Round((Convert.ToDecimal( dr.Cells["Cantidad"].Value * Convert.ToDecimal( dr.Cells["Precio"].Value) * Convert.ToDecimal( dr.Cells["DescuentoPorc"].Value/100, 2);

					dr.Cells["DescuentoArt"].Value = dDescuentoLineal;

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{
						this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
						this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
					else
					{
						this.txtIva0.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
						this.txtDescIva0Total.Value = Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal( dr.Cells["DescuentoArt"].Value), 2);
					}
				}

				this.txtBaseImponible.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value), 2);

				this.txtIvaTotal.Value = Math.Round((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * 12.00m /100.00m, 2);

				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 2) 
					- Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value), 2) 
					- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 2);
			}
			#endregion Desplazamiento de grilla			
		}

		private bool ValidaPrecio(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			
			decimal dPrecio = (decimal)e.Cell.Row.Cells["Precio"].Value;
			
			dPrecio = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", idTipoPrecio, (int)e.Cell.Row.Cells["idArticulo"].Value));			
	
			if ((decimal)e.Cell.Value < dPrecio) 
			{
				e.Cell.Value = dPrecio;
				MessageBox.Show(string.Format("No puede poner precio menor a {0}.", Math.Round(dPrecio, 2)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
		
			return true;
		}

		private bool ValidaDescuento(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			
			#region Validacion
			if (!miAcceso.CambiarDescuento) 
			{
				e.Cell.Value = (decimal)e.Cell.OriginalValue;
				MessageBox.Show("No tiene Acceso a Modificar Descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				return false;
			}
			
			if (Math.Round((decimal)e.Cell.Value, 2) > 100.00m)
			{
				e.Cell.Value = (decimal)e.Cell.OriginalValue;				
				return false;
			}

			if (!bDescuentoActivo)
			{
				//				decimal dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Descuento From CompraFormaPago Where idFormaPago = {0}", (int)this.cmbFormaPago.Value));
				//
				//				if (Math.Round((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value, 2) > Math.Round(dDescuento, 2))
				//				{
				//					e.Cell.Value = (decimal)e.Cell.OriginalValue;
				//					MessageBox.Show("El Descuento no puede ser mayor al permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				//					return false;
				//				}
			
//				if ((int)e.Cell.Row.Cells["idCombo"].Value == 0)  
//				{
				if (!bDescuentoActivo)  
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetallePromocion dp Inner Join CompraPromocion cp On cp.idPromocion = dp.idPromocion Where dp.idArticulo = {0} And Promocional = 1 And Estado = 'ACTIVO'", (int)e.Cell.Row.Cells["idArticulo"].Value)) > 0)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Producto Promocional No puede Aplicar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}
//				}

				if (!bDescuentoActivo)
				{
					if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 && !(bool)e.Cell.Row.Cells["Principal"].Value)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Producto Promocional No puede Modificar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}

				int idTPrecio = idTipoPrecio;
				if ((int)this.cmbFormaPago.Value == 1) idTPrecio = 2;
				if ((int)this.cmbFormaPago.Value == 11) idTPrecio = 90;

//				bool bDescuento = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaLimiteDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6})",
//					(int)this.cmbBodega.Value, (int)e.Cell.Row.Cells["idArticulo"].Value, idTPrecio, (decimal)e.Cell.Row.Cells["DescuentoPorc"].Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value), true);
				
//				if (!bDescuento)
//				{
//					e.Cell.Value = (decimal)e.Cell.OriginalValue;
//					MessageBox.Show("El Descuento debe estar en el Rango Permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
//					return false;
//				}

				int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
				if ((int)e.Cell.Row.Cells["idCombo"].Value > 0 && !(bool)e.Cell.Row.Cells["Principal"].Value && (int)e.Cell.Row.Cells["Tipo"].Value == 1)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value && ((int)e.Cell.Row.Cells["idCombo"].Value == (int)dr.Cells["idCombo"].Value)) 
						{
							idArticulo = (int)dr.Cells["idArticulo"].Value;
							continue;
						}
					}
				}

				decimal dDescuento = 0.00m;
				bool bBloquea = false;
				if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 && (bool)e.Cell.Row.Cells["Principal"].Value)
				{
					dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaDescuentoPromociones {0}, {1}", idArticulo, idTPrecio));
					bBloquea = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec RetornaEstadoDescuentoPromociones {0}", idArticulo), true);
				}
				else
				{
					dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaLimiteDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
						(int)this.cmbBodega.Value, idArticulo, idTPrecio, (decimal)e.Cell.Row.Cells["DescuentoPorc"].Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value));
				}

				decimal dSumaDescuento = 0.00m;
				
				if (!bBloquea)
				{
					dSumaDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaSumaDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6})",
						(int)this.cmbBodega.Value, idArticulo, idTPrecio, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value));
				}

				decimal dDescuentoPermitido = dDescuento + dSumaDescuento;

				if (dDescuentoPermitido > 100.00m) dDescuentoPermitido = 100.00m;

				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > dDescuentoPermitido)
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show("El Descuento debe estar en el Rango Permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
			}

			//			if ((int)e.Cell.Row.Cells["idCombo"].Value > 0)
			//			{
			//				if (!(bool)e.Cell.Row.Cells["Principal"].Value)
			//				{
			//					e.Cell.Value = (decimal)e.Cell.OriginalValue;	
			//					return false;
			//				}
			//			}

			//			if (Math.Round((decimal)e.Cell.Value, 2) > 5.00)
			//			{
			//				e.Cell.Value = 0.00;
			//				return false;
			//			}

			//			if ((decimal)e.Cell.Value > (decimal)e.Cell.OriginalValue) 
			//			{
			//				e.Cell.Value = (decimal)e.Cell.OriginalValue;
			//				MessageBox.Show(string.Format("Descuento Aprobado: {0}.", (decimal)e.Cell.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//				return false;
			//			}
			#endregion Validacion	
			return true;
		}

		private void InformacionArticulo(Infragistics.Win.UltraWinGrid.CellEventArgs e, int idArticulo)
		{			
			string sSQLInf = string.Format("Exec CargaInformacionArticulo {0}, {1}, {2}, {3}, {4}, {5}, {6}", idArticulo, (int)this.cmbFormaPago.Value, idTipoPrecio, (int)this.cmbBodega.Value, this.txtidBanco.Value, this.txtidTarjeta.Value, this.txtidPlan.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
			dr.Read();
			e.Cell.Row.Cells["idArticulo"].Value = idArticulo;
			e.Cell.Row.Cells["Codigo"].Value = dr.GetString(1);
			e.Cell.Row.Cells["Articulo"].Value = dr.GetString(2);
			e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(3);
			e.Cell.Row.Cells["Cantidad"].Value = 1;
			e.Cell.Row.Cells["Precio"].Value = Convert.ToDecimal(dr.GetValue(4));
			e.Cell.Row.Cells["Impuesto"].Value = Convert.ToDecimal(dr.GetValue(5));
			e.Cell.Row.Cells["DescuentoArt"].Value = Convert.ToDecimal(dr.GetValue(6));
			e.Cell.Row.Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.GetValue(7));
			e.Cell.Row.Cells["Tipo"].Value = 0;
			e.Cell.Row.Cells["Entregado"].Value = false;
			dr.Close();			
		}
		#endregion Funciones del grid

		private bool bExistencia(int idArticulo, bool bGrabar)
		{	
			if (!bConsignacion)
			{
				if ((int)this.cmbBodega.Value != 5)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", idArticulo), true))
					{
						DateTime dtFecha = (DateTime)this.dtFecha.Value;
						int iCantidad = 0;
						if (!bGrabar) iCantidad = 1;
						int iCantidadCompara = 0;
						int idArticuloCompara = 0;

						if (this.ultraGrid1.Rows.Count > 0)
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{
								idArticuloCompara = 0;

								if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticuloCompara = (int)dr.Cells["idArticulo"].Value;
					
								if (idArticulo == idArticuloCompara)
								{
									if (dr.Cells["Cantidad"].Value != System.DBNull.Value) iCantidadCompara = (int)dr.Cells["Cantidad"].Value;

									iCantidad = iCantidad + iCantidadCompara;
								}
							}
						}
						else iCantidad = 1;
			
						int Existencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.CalculaExistenciasInventario('{0}', {1}, {2})", dtFecha.ToString("yyyyMMdd"), idArticulo, (int)this.cmbBodega.Value));
						if (Existencia < iCantidad)
						{				
							MessageBox.Show(string.Format("Existencia Actual {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
							bEliminaAlValidar = true;
							if (!bGrabar)this.ultraGrid1.ActiveRow.Delete();
							bEliminaAlValidar = false;

							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
							return false;
						}	
						else return true;
					}
					else return true;
				}
				else return true;
			}
			else return true;
		}

		bool bArtPromocion = false;

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bEdicion) return;

			if (bCargaPromocionFP)
			{
				Total();
				return;
			}
			//			Total();

			bPromocionCombo = false;
			bPromocion = false;
			bCombo = false;
			int IdArticulo = 0;
			string stCod = "";
      
			//			bool Manual = false;
			#region Estado Factura manual
			//			if (e.Cell.Column.ToString() == "Manual")
			//			{
			//				int Cont = 0;
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			//				{
			//					Manual = (bool)dr.Cells["Manual"].Value;
			//					if (Manual) Cont = Cont + 1;
			//				}
			//
			//					if (Cont > 0)
			//				{
			//					drCompra["Estado"] = 11;
			//					this.lblEstado.Text = "Manual";
			//				}
			//				else if (Cont == 0)
			//				{
			//					drCompra["Estado"] = 1;
			//					this.lblEstado.Text = "";
			//				}				
			//
			//				if (!(bool)e.Cell.Row.Cells["Manual"].Value)
			//				{
			//					string sSQL = string.Format("Exec PointBuscaPorSerie '{0}', {1}", e.Cell.Row.Cells["SerieL"].Value.ToString(), (int) this.cmbBodega.Value);
			//					string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			//
			//					if (sCodigo.StartsWith("Error:")) 
			//					{
			//						//						MessageBox.Show("La serial existe, no puede crear como manual.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//						e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//						//return;
			//					}
			//				}
			//
			//				if (e.Cell.Row.Cells["SerieL"].Value != DBNull.Value)
			//				{
			//					if (!(bool)e.Cell.Row.Cells["Manual"].Value) return;
			//					string sSQL = string.Format("Exec PointBuscaPorSerie '{0}', {1}", e.Cell.Row.Cells["SerieL"].Value.ToString(), (int) this.cmbBodega.Value);
			//					string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			//
			//					if (!sCodigo.StartsWith("Error:")) 
			//					{
			//						MessageBox.Show("La serial existe, no puede crear como manual.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//						e.Cell.Row.Cells["Manual"].Value = false;
			//						return;
			//					}
			//				}
			//			}
			#endregion Estado Factura manual

			#region Busca por serial para venta
			//			if (IdTipoFactura == 1)
			//			{
			//				if (e.Cell.Column.ToString() == "SerieL")
			//				{
			//					if ((bool)e.Cell.Row.Cells["Manual"].Value)
			//					{
			//						if (e.Cell.Row.Cells["SerieL"].Value == DBNull.Value) return;
			//
			//						if (e.Cell.Row.Cells["idArticulo"].Value != DBNull.Value)
			//						{
			//							string stMensaje = string.Format("select isNull(ManejaSerial, 0) from Articulo where idArticulo = {0}", (int)e.Cell.Row.Cells["idArticulo"].Value);
			//							bool bManSer = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);
			//							if (!bManSer)
			//							{								
			//								MessageBox.Show(string.Format("El Artículo '{0}' no maneja serial, no puede crear como manual.", e.Cell.Row.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//								e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//								return;
			//							}
			//						}
			//						
			//						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Seriales Where Serial = '{0}'", e.Cell.Row.Cells["SerieL"].Value.ToString())) > 0)
			//						{
			//							MessageBox.Show(string.Format("La serial existe en el sistema.", e.Cell.Row.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//							e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//							return;
			//						}
			//
			//						string sSQL = string.Format("Exec PointBuscaPorSerie '{0}', {1}", e.Cell.Row.Cells["SerieL"].Value.ToString(), (int) this.cmbBodega.Value);
			//						string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			//
			//						if (!sCodigo.StartsWith("Error:")) 
			//						{
			//							MessageBox.Show("La serial existe, no puede crear como manual.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//							e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//							return;
			//						}
			//					}
			//					else
			//					{					
			//						if (e.Cell.Row.Cells["SerieL"].Value == DBNull.Value) return;
			//
			//						if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value)
			//						{
			//							string sSQL = string.Format("Exec PointBuscaPorSerie '{0}', {1}", e.Cell.Row.Cells["SerieL"].Value.ToString(), (int) this.cmbBodega.Value);
			//							string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			//
			//							if (!sCodigo.StartsWith("Error:")) 
			//							{
			//								e.Cell.Row.Cells["Codigo"].Value = sCodigo;
			//							}
			//							else 
			//							{
			//								MessageBox.Show(sCodigo.Substring(7), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//								e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//								return;
			//							}
			//						}
			//						else if (e.Cell.Row.Cells["idArticulo"].Value != DBNull.Value)
			//						{
			//							string stMensaje = string.Format("select isNull(ManejaSerial, 0) from Articulo where idArticulo={0}", (int)e.Cell.Row.Cells["idArticulo"].Value);
			//							bool bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);
			//							if (!bManejaSerial)
			//							{
			//								MessageBox.Show(string.Format("El Artículo '{0}' no maneja serial", e.Cell.Row.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//								e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//								return;
			//							}
			//
			//							string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}", e.Cell.Row.Cells["SerieL"].Value.ToString(), (int)e.Cell.Row.Cells["idArticulo"].Value, (int)this.cmbBodega.Value);
			//							int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			//        
			//							if (iCtaSerial < 1)
			//							{
			//								MessageBox.Show("La serial no existe o no pertenece al articulo seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//								e.Cell.Row.Cells["SerieL"].Value = DBNull.Value;
			//								return;
			//							}
			//						}
			//					}
			//				}
			//			}
			#endregion Busca por serial para venta

			if (e.Cell.Column.ToString() == "Codigo")
			{
				stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();
			}
			
			if (!bActualiza) return;
			
			bActualiza = false;
			#region Verifica Cambio
			int VerificaCambio = 0;
			#region Codigo o Articulo
			if (e.Cell.Column.ToString() == "Codigo" || e.Cell.Column.ToString() == "Articulo")
			{
				#region Variables
				int iCuenta = 0;
				int idPromocion = 0;
				int idCombo = 0;
				int idComboGrupo = 0;
				if (bRastreo) Rastreo("Inicio Cambio de " + e.Cell.Column.ToString());
				#endregion Variables
				#region Codigo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();
					#region cambia la comilla por guion
					if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRCxG'")) 				
					{						
						string cadena = revisaCadena(stCod);
						stCod = cadena;												
					}
					#endregion cambia la comilla por guion
				}
				#endregion Codigo
				#region Articulo
				if (e.Cell.Column.ToString() == "Articulo")
				{
					stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
				}
				#endregion Articulo

				#region Cuenta coincidencias del codigo
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, {3}, 0", stCod, (int)this.txtIdTipoFactura.Value, (int) this.cmbBodega.Value, 0);
				if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				
				if (iCuenta != 1)
				{ 
					#region Mas de una coincidencia o ninguna muestra la lista
					if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));

					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							if (bRastreo) Rastreo("Articulo escogido de la lista");
							IdArticulo = miBusqueda.idArticulo;

							#region Valida Combos - Promociones - Producto
							int idTPrecioPromoEsp = idTipoPrecio;
							if ((int)this.cmbFormaPago.Value == 1) idTPrecioPromoEsp = 2;
							idPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionEspecifica ({0}, {1}, {2}, {3}, {4}, {5}, {6})", 
								IdArticulo, idTPrecioPromoEsp, (int)this.cmbBodega.Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value));
							if (idPromocion > 0)
							{	
								#region Promocion
								int iFila = e.Cell.Row.Index;
								int idArtProm = 0;
								string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", idPromocion, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
								SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
								
								while (drC.Read())
								{
									idArtProm = drC.GetInt32(0);
									string sCodProm = drC.GetString(1); 
									string sArtProm = drC.GetString(2);
									int iUniProm = drC.GetInt32(3);
									decimal dPrecProm = drC.GetDecimal(4);
									decimal dImpProm = drC.GetDecimal(5);
									decimal dDesArtProm = drC.GetDecimal(6);
									decimal dDesPorProm = drC.GetDecimal(7);
									bool bPrincipal = drC.GetBoolean(8);
																	
									ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
									ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
									ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
									ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
									ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
									ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
									ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
									ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
									ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;									
									ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
									ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 2;
									ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocion;
									ultraGrid1.Rows[iFila].Cells["Entregado"].Value = false;
																		
									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
									iFila = iFila + 1;
								} 
								drC.Close();
								#endregion Promocion
							}					
							else
							{
								#region Valida Combos por Artículo
								if ((int)this.cmbBodega.Value != 1)
								{
									if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Combo Where idArticuloCombo = {0} And Estado = 0", IdArticulo)) > 0)
									{
										if (!bExistencia(IdArticulo, false)) 
										{
											if (bRastreo) Rastreo("Descuento menor al establecido");					
											bActualiza = true;
											Total();
											return;	
										}

										if (DialogResult.Yes == MessageBox.Show(string.Format("Este Artículo esta asociado a un combo. ¿Cargar los artículos del combo?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
										{											
											#region Combo
											idCombo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCombo From Combo Where idArticuloCombo = {0} And Estado = 0", IdArticulo));
											
											string sSQL = string.Format("Exec ConsultaExistenciaCombos {0}, {1}", idCombo, (int)this.cmbBodega.Value);
											if (Funcion.dvProcedimiento(cdsSeteoF, sSQL).Count > 0)
											{
												using (frmArticulosSinExistenciaEnCombos ASEC = new frmArticulosSinExistenciaEnCombos(idCombo, (int)this.cmbBodega.Value))
												{				
													if (DialogResult.No == ASEC.ShowDialog())
													{
														return;
													}
												}
											}

											int iFila = e.Cell.Row.Index;
											int idArtCom = 0;
											string sSQLProm = string.Format("Exec ArticulosCombos {0}, {1}, {2}, {3}, 1", (int)this.cmbFormaPago.Value, idCombo, idTipoPrecio, (int)this.cmbBodega.Value);
											SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);

											while (drC.Read())
											{									
												idArtCom = drC.GetInt32(0);
												string sCodCom = drC.GetString(1); 
												string sArtCom = drC.GetString(2);
												int iUniCom = drC.GetInt32(3);
												decimal dPrecCom = drC.GetDecimal(4);
												decimal dImpCom = drC.GetDecimal(5);
												decimal dDesArtCom = drC.GetDecimal(6);
												decimal dDesPorCom = drC.GetDecimal(7);									
												bool bPrincipal = drC.GetBoolean(8);
												bool bBloqueado = drC.GetBoolean(9);

												ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtCom;
												ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodCom;
												ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtCom;
												ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniCom;
												ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
												ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecCom;
												ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpCom;
												ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtCom;
												ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorCom;									
												ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
												ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 1;
												ultraGrid1.Rows[iFila].Cells["Entregado"].Value = false;
												ultraGrid1.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
												ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idCombo;
																					
												this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
												iFila = iFila + 1;
											} 
											drC.Close();
											#endregion Combo 							
										}
										else
										{
											#region Asignaciones si no quiere el combo
											if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
					
											if (!bExistencia(IdArticulo, false)) 
											{
												if (bRastreo) Rastreo("Descuento menor al establecido");					
												bActualiza = true;
												Total();
												return;	
											}
											
											InformacionArticulo(e, IdArticulo);
						
											VerificaCambio = 2; // Verificaciones y Cambio de precio
											#endregion Asignaciones si no quiere el combo
										}
									}
									else if(Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ComboPorSubGrupo {0}", IdArticulo)) > 0)
									{
										if (!bExistencia(IdArticulo, false)) 
										{
											if (bRastreo) Rastreo("NO hay Existencia");
											bActualiza = true;
											Total();
											return;
										}

										idComboGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ComboPorSubGrupo {0}", IdArticulo));
										if (DialogResult.Yes == MessageBox.Show(string.Format("Este Artículo esta asociado a un combo. ¿Cargar los artículos del combo?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
										{
											if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo del Combo por Subgrupo");
											
											string sSQL = string.Format("Exec ConsultaExistenciaCombos {0}, {1}", idComboGrupo, (int)this.cmbBodega.Value);
											if (Funcion.dvProcedimiento(cdsSeteoF, sSQL).Count > 0)
											{
												using (frmArticulosSinExistenciaEnCombos ASEC = new frmArticulosSinExistenciaEnCombos(idComboGrupo, (int)this.cmbBodega.Value))
												{				
													if (DialogResult.No == ASEC.ShowDialog())
													{
														return;
													}
												}
											}

											InformacionArticulo(e, IdArticulo);
											
											e.Cell.Row.Cells["idCombo"].Value = idComboGrupo;
											e.Cell.Row.Cells["Principal"].Value = true;
											
											#region Combo por SubGrupo
											int iFila = e.Cell.Row.Index;
											int idArtCom = 0;
											string sSQLProm = string.Format("Exec ArticulosCombos {0}, {1}, {2}, {3}, 1", this.cmbFormaPago.Value, idComboGrupo, idTipoPrecio, (int)this.cmbBodega.Value);
											SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);

											while (drC.Read())
											{
												this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
												iFila = iFila + 1;
												idArtCom = drC.GetInt32(0);
												string sCodCom = drC.GetString(1); 
												string sArtCom = drC.GetString(2);
												int iUniCom = drC.GetInt32(3);
												decimal dPrecCom = drC.GetDecimal(4);
												decimal dImpCom = drC.GetDecimal(5);
												decimal dDesArtCom = drC.GetDecimal(6);
												decimal dDesPorCom = drC.GetDecimal(7);
												bool bPrincipal = drC.GetBoolean(8);
												bool bBloqueado = drC.GetBoolean(9);

												ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtCom;
												ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodCom;
												ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtCom;
												ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniCom;
												ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
												ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecCom;
												ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpCom;
												ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtCom;
												ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorCom;
												ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
												ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 1;
												ultraGrid1.Rows[iFila].Cells["Entregado"].Value = false;
												ultraGrid1.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
												ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idComboGrupo;											
											} 
											drC.Close();

											this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
											#endregion Combo por SubGrupo 
										}
										else
										{
											#region Asignaciones si no existe el combo
											if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

											if (!bExistencia(IdArticulo, false))
											{
												if (bRastreo) Rastreo("No hay Existencia");
												bActualiza = true;
												Total();
												return;
											}
											
											InformacionArticulo(e, IdArticulo);

											VerificaCambio = 2; // Verificaciones y Cambio de precio
											#endregion Asignaciones si no existe el combo
										}
									}
									else
									{
										#region Asignaciones si no existe el combo
										if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

										if (!bExistencia(IdArticulo, false))
										{
											if (bRastreo) Rastreo("No hay Existencia");
											bActualiza = true;
											Total();
											return;
										}
										
										InformacionArticulo(e, IdArticulo);

										VerificaCambio = 2; // Verificaciones y Cambio de precio
										#endregion Asignaciones si no existe el combo
									}
								}
								else
								{
									#region Asignaciones si es Bodega Principal
									if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

									if (!bExistencia(IdArticulo, false))
									{
										if (bRastreo) Rastreo("No hay Existencia");
										bActualiza = true;
										Total();
										return;
									}

									InformacionArticulo(e, IdArticulo);

									VerificaCambio = 2; // Verificaciones y Cambio de precio
									#endregion Asignaciones si es Bodega Principal
								}
								#endregion Valida Combos por Artículo
							}

							this.Total();

							#region Promocion Por Valor
							if (!bArtPromocion)
							{
								int idPromocionValor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, 0, 0, 0, 0)", IdArticulo, idTipoPrecio, (int)this.cmbBodega.Value, this.txtTotal.Value));
								if (idPromocionValor > 0 && (int)this.cmbBodega.Value != 1)
								{									
									int iFila = 0;
									if (idCombo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
									else if (idComboGrupo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
									else if (idPromocion > 0) iFila = this.ultraGrid1.Rows.Count - 1;
									else 
									{
										iFila = e.Cell.Row.Index + 1;
										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
									}
								
									int idArtProm = 0;
									string sSQLProm = string.Format("Exec CargaArticulosPromocionPorValor {0}, {1}, {2}, {3}, {4}", idPromocionValor, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value, this.SumaPrecios());
									SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
									
									#region Productos Promociones
									while (drC.Read())
									{										
										idArtProm = drC.GetInt32(0);
										string sCodProm = drC.GetString(1); 
										string sArtProm = drC.GetString(2);
										int iUniProm = drC.GetInt32(3);
										decimal dPrecProm = drC.GetDecimal(4);
										decimal dImpProm = drC.GetDecimal(5);
										decimal dDesArtProm = drC.GetDecimal(6);
										decimal dDesPorProm = drC.GetDecimal(7);									
										bool bPrincipal = drC.GetBoolean(8);

										bool bValArtCombo = false;

										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
										{
											int idArtPromCarg = 0;
											if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
											if (idArtProm == idArtPromCarg) bValArtCombo = true;
										}

										if (!bValArtCombo)
										{
											ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
											ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
											ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
											ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
											ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
											ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
											ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
											ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
											ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
											ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
											ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocionValor;
											ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 3;

											this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

											iFila = iFila + 1;
										}
									} 
									drC.Close();
									#endregion Productos Promociones

									bArtPromocion = true;
								}
							}
							#endregion Promocion Por Valor

							#region Promocion Adicional
							int idPromocionAdicional = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionAdicional ({0}, {1}, {2}, 0, 0, 0, 0)", (int)this.cmbBodega.Value, IdArticulo, idTipoPrecio));
							if (idPromocionAdicional > 0 && (int)this.cmbBodega.Value != 1)
							{	
								int iFila = 0;
								if (idCombo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else if (idComboGrupo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else if (idPromocion > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else 
								{
									iFila = e.Cell.Row.Index + 1;
									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();									
								}
								
								int idArtProm = 0;
								
								string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", idPromocionAdicional, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
								SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
							
								#region Productos Promociones
								while (drC.Read())
								{
									idArtProm = drC.GetInt32(0);
									string sCodProm = drC.GetString(1); 
									string sArtProm = drC.GetString(2);
									int iUniProm = drC.GetInt32(3);
									decimal dPrecProm = drC.GetDecimal(4);
									decimal dImpProm = drC.GetDecimal(5);
									decimal dDesArtProm = drC.GetDecimal(6);
									decimal dDesPorProm = drC.GetDecimal(7);									
									bool bPrincipal = drC.GetBoolean(8);
								
									bool bValArtCombo = false;

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
									{
										int idArtPromCarg = 0;
										if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
										if (idArtProm == idArtPromCarg) bValArtCombo = true;
									}

									if (!bValArtCombo)
									{
										ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
										ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
										ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
										ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
										ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
										ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
										ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
										ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
										ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
										ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
										ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocionAdicional;
										ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 3;

										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

										iFila = iFila + 1;
									}
								} 
								drC.Close();						
								#endregion Productos Promociones

								bArtPromocion = true;
							}					
							#endregion Promocion Adicional
							#endregion Valida Combos - Promociones - Producto
						}
						else
						{
							if (bRastreo) Rastreo("Cancela la selección de Artículos en lista");
							bActualiza = true;
							return;						
						}
					} 
					#endregion Mas de una coincidencia o ninguna muestra la lista
				}
				else 
				{
					#region Encuentra un solo articulo 
					if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
					string sSQLArtEnc = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value);
					IdArticulo = Funcion.iEscalarSQL(cdsSeteoF, sSQLArtEnc, true);						
					#endregion Encuentra un solo articulo
					
					if (!bModificaColor)
					{						
						int idTPrecioPromoEsp = idTipoPrecio;
						if ((int)this.cmbFormaPago.Value == 1 || (int)this.cmbFormaPago.Value == 11) idTPrecioPromoEsp = 2;
						idPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionEspecifica ({0}, {1}, {2}, {3}, {4}, {5}, {6})", 
							IdArticulo, idTPrecioPromoEsp, (int)this.cmbBodega.Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value));
						if (idPromocion > 0 && (int)this.cmbBodega.Value != 1)
						{	
							#region Promocion
							int iFila = e.Cell.Row.Index;
							int idArtProm = 0;
							string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", idPromocion, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
							SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
							
							while (drC.Read())
							{
								idArtProm = drC.GetInt32(0);
								string sCodProm = drC.GetString(1); 
								string sArtProm = drC.GetString(2);
								int iUniProm = drC.GetInt32(3);
								decimal dPrecProm = drC.GetDecimal(4);
								decimal dImpProm = drC.GetDecimal(5);
								decimal dDesArtProm = drC.GetDecimal(6);
								decimal dDesPorProm = drC.GetDecimal(7);									
								bool bPrincipal = drC.GetBoolean(8);

								ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
								ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
								ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
								ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
								ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
								ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
								ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
								ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
								ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
								ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
								ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 2;
								ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocion;

								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

								iFila = iFila + 1;
							} 
							drC.Close();
							#endregion Promocion
						}
						else
						{
							#region Valida Combos por Artículo
							if ((int)this.cmbBodega.Value != 1)
							{
								if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Combo Where idArticuloCombo = {0} And Estado = 0", IdArticulo)) > 0)
								{
									if (!bExistencia(IdArticulo, false)) 
									{
										if (bRastreo) Rastreo("Descuento menor al establecido");					
										bActualiza = true;
										Total();
										return;	
									}

									if (DialogResult.Yes == MessageBox.Show(string.Format("Este Artículo esta Asociado a un Combo\n\n¿Cargar los Artículos del Combo?"), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
									{									
										#region Combo
										idCombo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCombo From Combo Where idArticuloCombo = {0} And Estado = 0", IdArticulo));

										string sSQL = string.Format("Exec ConsultaExistenciaCombos {0}, {1}", idCombo, (int)this.cmbBodega.Value);
										if (Funcion.dvProcedimiento(cdsSeteoF, sSQL).Count > 0)
										{
											using (frmArticulosSinExistenciaEnCombos ASEC = new frmArticulosSinExistenciaEnCombos(idCombo, (int)this.cmbBodega.Value))
											{				
												if (DialogResult.No == ASEC.ShowDialog())
												{
													return;
												}
											}
										}

										int iFila = e.Cell.Row.Index;
										int idArtCom = 0;
										string sSQLProm = string.Format("Exec ArticulosCombos {0}, {1}, {2}, {3}, 1", (int)this.cmbFormaPago.Value, idCombo, idTipoPrecio, (int)this.cmbBodega.Value);
										SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);

										while (drC.Read())
										{	
											//									this.ultraGrid1.Rows.Band.AddNew();
											idArtCom = drC.GetInt32(0);
											string sCodCom = drC.GetString(1); 
											string sArtCom = drC.GetString(2);
											int iUniCom = drC.GetInt32(3);
											decimal dPrecCom = drC.GetDecimal(4);
											decimal dImpCom = drC.GetDecimal(5);
											decimal dDesArtCom = drC.GetDecimal(6);
											decimal dDesPorCom = drC.GetDecimal(7);									
											bool bPrincipal = drC.GetBoolean(8);
											bool bBloqueado = drC.GetBoolean(9);

											ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtCom;
											ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodCom;
											ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtCom;
											ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniCom;
											ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
											ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecCom;
											ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpCom;
											ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtCom;
											ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorCom;
											ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
											ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 1;
											ultraGrid1.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
											ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idCombo;
									
											this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
									
											iFila = iFila + 1;									
										}
										drC.Close();
								
										//this.ultraGrid1.Rows.Band.AddNew();
										//this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
										#endregion Combo 																																															
									}
									else
									{
										#region Asignaciones si no quiere el combo
										if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
					
										if (!bExistencia(IdArticulo, false))
										{
											if (bRastreo) Rastreo("No hay Existencia");					
											bActualiza = true;
											Total();
											return;
										}

										InformacionArticulo(e, IdArticulo);
						
										VerificaCambio = 2; // Verificaciones y Cambio de precio
										#endregion Asignaciones si no quiere el combo
									}
								}
								else if(Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ComboPorSubGrupo {0}", IdArticulo)) > 0)
								{
									if (!bExistencia(IdArticulo, false))
									{
										if (bRastreo) Rastreo("No Hay Existencia");					
										bActualiza = true;
										Total();
										return;
									}

									idComboGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ComboPorSubGrupo {0}", IdArticulo));
									if (DialogResult.Yes == MessageBox.Show(string.Format("Este Artículo esta asociado a un combo. ¿Cargar los artículos del combo?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
									{
										string sSQL = string.Format("Exec ConsultaExistenciaCombos {0}, {1}", idComboGrupo, (int)this.cmbBodega.Value);
										if (Funcion.dvProcedimiento(cdsSeteoF, sSQL).Count > 0)
										{
											using (frmArticulosSinExistenciaEnCombos ASEC = new frmArticulosSinExistenciaEnCombos(idComboGrupo, (int)this.cmbBodega.Value))
											{				
												if (DialogResult.No == ASEC.ShowDialog())
												{
													return;
												}
											}
										}

										if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo del Combo por Subgrupo");
																					
										InformacionArticulo(e, IdArticulo);
										
										e.Cell.Row.Cells["idCombo"].Value = idComboGrupo;
										e.Cell.Row.Cells["Principal"].Value = true;
										#region Combo por SubGrupo
										int iFila = e.Cell.Row.Index;
										int idArtCom = 0;
										string sSQLProm = string.Format("Exec ArticulosCombos {0}, {1}, {2}, {3}, 1", (int)this.cmbFormaPago.Value, idComboGrupo, idTipoPrecio, (int)this.cmbBodega.Value);
										SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);

										while (drC.Read())
										{
											this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
											iFila = iFila + 1;
											idArtCom = drC.GetInt32(0);
											string sCodCom = drC.GetString(1); 
											string sArtCom = drC.GetString(2);
											int iUniCom = drC.GetInt32(3);
											decimal dPrecCom = drC.GetDecimal(4);
											decimal dImpCom = drC.GetDecimal(5);
											decimal dDesArtCom = drC.GetDecimal(6);
											decimal dDesPorCom = drC.GetDecimal(7);									
											bool bPrincipal = drC.GetBoolean(8);
											bool bBloqueado = drC.GetBoolean(9);

											ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtCom;
											ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodCom;
											ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtCom;
											ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniCom;
											ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
											ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecCom;
											ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpCom;
											ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtCom;
											ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorCom;
											ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
											ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 1;
											ultraGrid1.Rows[iFila].Cells["Bloqueado"].Value = bBloqueado;
											ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idComboGrupo;											
										} 
										drC.Close();
									
										iFila = iFila + 1;
	
										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
									
										string sSQLVPS = string.Format("Select dbo.ValidaPromocionSubGrupo({0}, {1}, {2})", (int)this.cmbBodega.Value, IdArticulo, idTipoPrecio);		
										int IPromoSubGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQLVPS);

										if (IPromoSubGrupo > 0)
										{
											#region Promocion por Subgrupo
											int idArtPromSG = 0;
											string sSQLPromSG = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", IPromoSubGrupo, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
											SqlDataReader drSG = Funcion.rEscalarSQL(cdsSeteoF, sSQLPromSG, true);

											while (drSG.Read())
											{
												idArtPromSG = drSG.GetInt32(0);
												string sCodProm = drSG.GetString(1); 
												string sArtProm = drSG.GetString(2);
												int iUniProm = drSG.GetInt32(3);
												decimal dPrecProm = drSG.GetDecimal(4);
												decimal dImpProm = drSG.GetDecimal(5);
												decimal dDesArtProm = drSG.GetDecimal(6);
												decimal dDesPorProm = drSG.GetDecimal(7);									
												bool bPrincipal = drSG.GetBoolean(8);

												ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtPromSG;
												ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
												ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
												ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
												ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
												ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
												ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
												ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
												ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
												ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
												ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 2;

												this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

												iFila = iFila + 1;
											} 
											drSG.Close();
											#endregion Promocion por Subgrupo
										}

										//this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

										#endregion Combo por SubGrupo 
									}
									else
									{
										#region Asignaciones si no existe el combo
										if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

										if (!bExistencia(IdArticulo, false))
										{
											if (bRastreo) Rastreo("No hay Existencia");
											bActualiza = true;
											Total();
											return;
										}
										
										InformacionArticulo(e, IdArticulo);

										VerificaCambio = 2; // Verificaciones y Cambio de precio
										#endregion Asignaciones si no existe el combo
									}
								}
								else
								{
									int ValIdPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idPromocion Int = 0 If Exists(Select idPromocion From CompraPromocion Where idArticulo = {0}) Select @idPromocion = idPromocion From CompraPromocion Where idArticulo = {0} And Estado = 'FINALIZADO' Select @idPromocion", IdArticulo));
									if (ValIdPromocion > 0)
									{
										e.Cell.Row.Cells["Codigo"].Value = "";
										MessageBox.Show("Esta Promoción esta Caducada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
										return;
									}

									ValIdPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idPromocion Int = 0 If Exists(Select idPromocion From CompraPromocion Where idArticulo = {0}) Select @idPromocion = idPromocion From CompraPromocion Where idArticulo = {0} Select @idPromocion", IdArticulo));
									if (ValIdPromocion > 0)
									{
										e.Cell.Row.Cells["Codigo"].Value = "";
										MessageBox.Show("Esta Promoción No Cumple con los parametros necesarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										return;
									}

									#region Asignaciones si no existe el combo
									if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

									if (!bExistencia(IdArticulo, false))
									{
										if (bRastreo) Rastreo("No hay Existencia");
										bActualiza = true;
										Total();
										return;
									}
									
									InformacionArticulo(e, IdArticulo);

									VerificaCambio = 2; // Verificaciones y Cambio de precio
									#endregion Asignaciones si no existe el combo
								}
							}
							else
							{
								int ValIdPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idPromocion Int = 0 If Exists(Select idPromocion From CompraPromocion Where idArticulo = {0}) Select @idPromocion = idPromocion From CompraPromocion Where idArticulo = {0} Select @idPromocion", IdArticulo));
								if (ValIdPromocion > 0)
								{
									e.Cell.Row.Cells["Codigo"].Value = "";
									MessageBox.Show("Esta Promoción esta caducada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
									return;
								}

								ValIdPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idPromocion Int = 0 If Exists(Select idPromocion From CompraPromocion Where idArticulo = {0}) Select @idPromocion = idPromocion From CompraPromocion Where idArticulo = {0} Select @idPromocion", IdArticulo));
								if (ValIdPromocion > 0)
								{
									e.Cell.Row.Cells["Codigo"].Value = "";
									MessageBox.Show("Esta Promoción No Cumple con los parametros necesarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									return;
								}

								#region Asignaciones si es Bodega Principal
								if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

								if (!bExistencia(IdArticulo, false))
								{
									if (bRastreo) Rastreo("No hay Existencia");					
									bActualiza = true;
									Total();
									return;
								}
								
								InformacionArticulo(e, IdArticulo);

								VerificaCambio = 2; // Verificaciones y Cambio de precio
								#endregion Asignaciones si es Bodega Principal
							}	
							#endregion Valida Combos por Artículo
						}
						
						this.Total();

						#region Promocion Por Valor
//						if (!bArtPromocion)
//						{
							int idPromocionValor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, 0, 0, 0, 0)", 
								(int)this.cmbBodega.Value, IdArticulo, idTipoPrecio, this.txtTotal.Value));
							if (idPromocionValor > 0 && (int)this.cmbBodega.Value != 1)
							{	
								int iFila = 0;								
								if (idCombo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else if (idComboGrupo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else if (idPromocion > 0) iFila = this.ultraGrid1.Rows.Count - 1;
								else 
								{
									iFila = e.Cell.Row.Index + 1;
									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
								}
								
								if (iFila == 0) 
								{
									iFila = 1;
									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
								}

								int idArtProm = 0;
								
								string sSQLProm = string.Format("Exec CargaArticulosPromocionPorValor {0}, {1}, {2}, {3}, {4}", idPromocionValor, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value, this.SumaPrecios());
								SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
							
								#region Productos Promociones
								while (drC.Read())
								{
									idArtProm = drC.GetInt32(0);
									string sCodProm = drC.GetString(1); 
									string sArtProm = drC.GetString(2);
									int iUniProm = drC.GetInt32(3);
									decimal dPrecProm = drC.GetDecimal(4);
									decimal dImpProm = drC.GetDecimal(5);
									decimal dDesArtProm = drC.GetDecimal(6);
									decimal dDesPorProm = drC.GetDecimal(7);									
									bool bPrincipal = drC.GetBoolean(8);
								
									bool bValArtCombo = false;

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
									{
										int idArtPromCarg = 0;
										if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
										if (idArtProm == idArtPromCarg) bValArtCombo = true;
									}

									if (!bValArtCombo)
									{
										ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
										ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
										ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
										ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
										ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
										ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
										ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
										ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
										ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
										ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
										ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocionValor;
										ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 3;

										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

										iFila = iFila + 1;
									}
								} 
								drC.Close();						
								#endregion Productos Promociones

								bArtPromocion = true;
							}
//						}
						#endregion Promocion Por Valor

						#region Promocion Adicional
						int idPromocionAdicional = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionAdicional ({0}, {1}, {2}, 0, 0, 0, 0)", (int)this.cmbBodega.Value, IdArticulo, idTipoPrecio));
						if (idPromocionAdicional > 0 && (int)this.cmbBodega.Value != 1)
						{	
							int iFila = 0;							
							if (idCombo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
							else if (idComboGrupo > 0) iFila = this.ultraGrid1.Rows.Count - 1;
							else if (idPromocion > 0) iFila = this.ultraGrid1.Rows.Count - 1;
							else 
							{
								iFila = e.Cell.Row.Index + 1;
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();									
							}
							
							if (iFila == 0) 
							{
								iFila = 1;
								this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
							}

							int idArtProm = 0;
								
							string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", idPromocionAdicional, idTipoPrecio, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
							SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);
							
							#region Productos Promociones
							while (drC.Read())
							{
								idArtProm = drC.GetInt32(0);
								string sCodProm = drC.GetString(1); 
								string sArtProm = drC.GetString(2);
								int iUniProm = drC.GetInt32(3);
								decimal dPrecProm = drC.GetDecimal(4);
								decimal dImpProm = drC.GetDecimal(5);
								decimal dDesArtProm = drC.GetDecimal(6);
								decimal dDesPorProm = drC.GetDecimal(7);									
								bool bPrincipal = drC.GetBoolean(8);
								
								bool bValArtCombo = false;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
								{
									int idArtPromCarg = 0;
									if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
									if (idArtProm == idArtPromCarg) bValArtCombo = true;
								}

								if (!bValArtCombo)
								{
									ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = idArtProm;
									ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodProm;
									ultraGrid1.Rows[iFila].Cells["Articulo"].Value = sArtProm;
									ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
									ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
									ultraGrid1.Rows[iFila].Cells["Precio"].Value = dPrecProm;
									ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
									ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
									ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
									ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;
									ultraGrid1.Rows[iFila].Cells["idCombo"].Value = idPromocionAdicional;
									ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 3;

									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

									iFila = iFila + 1;
								}
							} 
							drC.Close();						
							#endregion Productos Promociones

							bArtPromocion = true;
						}					
						#endregion Promocion Adicional
					}
				}
			} 
			#endregion Codigo o articulo

			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
				VerificaCambio = 1; 
				if (!bConsignacion)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
					{
						if ((bool)e.Cell.Row.Cells["Principal"].Value) 
						{
							e.Cell.Value = 1;
							MessageBox.Show("Para Productos Principales de un Combo o Promocion la Cantidad es 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}

						if ((int)e.Cell.Row.Cells["Tipo"].Value > 0) 
						{
							e.Cell.Value = 1;
							MessageBox.Show("Para Productos que forman parte de un Combo o Promocion la Cantidad es 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}

						DateTime dtFecha = (DateTime)this.dtFecha.Value;
						int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
						int Existencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.CalculaExistenciasInventario('{0}', {1}, {2})", dtFecha.ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idArticulo"].Value, (int)this.cmbBodega.Value));
						int iCantidad = 0;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
						}

						if (iCantidad > Existencia)
						{					
							if (Existencia < 1)
							{
								bEliminaAlValidar = true;
								e.Cell.Row.Delete();
								bEliminaAlValidar = false;
								MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}
							else
							{
								e.Cell.Value = 1;
								MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}				
						}
					}
				}
				Total();
			}
			#endregion Cantidad

			if (e.Cell.Column.ToString() == "Precio")
			{				
				if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value) e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
				if (!miAcceso.CambiarPrecio)
				{
					if (!ValidaPrecio(e)) 
					{						
						if (bRastreo) Rastreo("Precio menor al establecido");
						Total();
						bActualiza = true;
						return;
					}
				}	
				Total();
			}

			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{		
				if (e.Cell.Row.Cells["DescuentoPorc"].Value == System.DBNull.Value) e.Cell.Row.Cells["DescuentoPorc"].Value = (decimal)e.Cell.OriginalValue;
				
				bActualiza = false;
					
				if (!ValidaDescuento(e))
				{
					if (bRastreo) Rastreo("Descuento menor al establecido");					
					bActualiza = true;
					Total();
					return;
				}
					
				Total();

				bActualiza = true;									
			}
			#endregion Fin de Verificacion de Cambio

			Total();

			bActualiza = true;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bEdicion = false;
			bNuevo = false;

			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			
			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtComprobante.Text = "";
			this.txtNumeroCredito.Text = "";			
			this.lblEstado.Text = "";

			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today.AddDays(-365);
			this.dtFechaRetencion.Value = System.DBNull.Value;

			this.txtNumEstado.Value = 0;
			this.txtNumIdAsiento.Value = 0;
			this.txtNumIdCliente.Value = 0;
			this.txtNumIdCompra.Value = 0;
			this.txtNumidFPNC.Value = 0;
			this.txtNumidOrigen.Value = 0;
			this.txtNumidOrigen1.Value = 0;
			this.txtNumIdProyecto.Value = 0;
			this.txtNumContadoCredito.Value = 0;
			this.txtNumIdproyectoZoom.Value = 0;
			this.txtidRetencion.Value = 0;			
			this.txtIdAsientoRetencion.Value = 0;
			this.txtIdEF.Value = 0;
			this.txtidBanco.Value = 0;
			this.txtidTarjeta.Value = 0;
			this.txtidPlan.Value = 0;
			this.txtidPlazo.Value = 0;
			this.txtCuotaDeEntrada.Value = 0.00m;			
			this.lblUsuario.Text = "";
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtBaseImponible.Value = 0.00m;

			dCuotaAsignada = 0.00m;
			dInteres = 0.00m;
			dCuotaDisponible = 0.00m;
			dCuotaUsada = 0.00m;
			AutorizaEntrada = false;
			dValorCuota = 0.00m;

			this.cmbFormaPago.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbTipoRuc.Enabled = false;

			this.txtNumero.Enabled = false;
			this.txtRuc.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumeroCredito.Enabled = false;
			this.dtFechaRetencion.Enabled = false;
						
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnTemporal.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Text = "Anular";
			this.btnCliente.Enabled = false;

			this.lblPlan.Text = "";

			if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;

			this.btnImprimeDctos.Enabled = false;
			this.btnClaveDescuento.Enabled = false;

			ModificaGrilla(false);
			HabilitaTab(false);
			this.grdSerial.Enabled = true;
			ModificaGrillaSerial(false);

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");

			#region Tab
			#region Tab Seriales
			this.txtSerial.Text = "";
			this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec RefrescaSerialVenta 0");	
			#endregion Tab Seriales
			#region Tab Retenciones
			this.txtSerieRetencion.Text = "";
			this.txtAutorizacionRetencion.Text = "";
			this.txtNumeroRetencion.Text = "";
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaRetDetalle - 1");
			this.txtSerieRetencion.Enabled = false;
			this.txtNumeroRetencion.Enabled = false;
			this.txtAutorizacionRetencion.Enabled = false;
			this.btnRetener.Enabled = false;
			this.btnCancelarRetencion.Enabled = false;
			#endregion Tab Retenciones
			#region Tab S.R.I.
			this.txtSerie.Text = "";
			this.txtAutFactura.Text = "";
			this.cmbCaducidad.Value = System.DBNull.Value;
			#endregion Tab S.R.I.
			#region Tab Entrega
			this.chkEntregado.Checked = false;
			this.cmbDirecciones.Value = System.DBNull.Value;
			this.dtFechaEntrega.Value = System.DBNull.Value;
			this.dtpHoraEntrega.Value = System.DBNull.Value;
			#endregion Tab Entrega
			#region Tab Nota
			this.txtNotas.Text = "";	
			#endregion Tab Nota
			#endregion Tab

			#region Variables
			bFormaPagoPromocional = false;
			bCargaPromocionFP = false;
			idTipoPrecio = 0;
			bNumeracionAutomatica = false;
			bImpreso = false;			
			BUsoNotaDeCredito = false;
			dSaldoNotaDeCredito = 0.00m;
			BRetencion = false;
			ModSerial = false;
			bEliminaAlValidar = false;			
			dAnticipo = 0.00m;
			bUsoAnticipo = false;
			bPromocionCombo = false;
			bPromocion = false;
			bCombo = false;
			bConsignacion = false;			
			bRetencion = false;
			dValorRetencion = 0.00m; 
			bAnticipo = false;
			dValorAnticipo = 0.00m;
			bCrediPoint = false;
			dValorCrediPoint = 0.00m;
			dEntrada = 0.00m;
			dTotalCuotas = 0.00m;
			NumCuotas = 0;
			dtFechPrimPago = DateTime.Today;
			bAnulado = false;
			idTipoOrigen = 0;
			bDescuentoActivo = false;
			bModificaColor = false;
			bArtPromocion = false;
			idBloqueaTransacciones = 0;
			dtFecVencCupon = DateTime.Today;
			sNumero = "";
			#endregion Variables

			this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieL"].CellActivation = Activation.AllowEdit;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.AllowEdit;
			//this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = Activation.AllowEdit;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = Activation.AllowEdit;

			this.btnNuevo.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{						
			bool Manual = false;
			string sSQL = "";
			decimal dRetenido = 0;
			decimal dRetenidoIVA = 0;
			decimal dRetenidoRenta = 0;
			int idDetFormaPagoRetencion = 0;
			DateTime dtFechaCadR = DateTime.Today;

			#region Validacion
			#region Vendedor
			if (this.cmbVendedor.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbVendedor.Focus();
				return;
			}
			if (this.cmbVendedor.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbVendedor.Focus();
				return;
			}
			#endregion Vendedor
			#region Numero
			if (!bNumeracionAutomatica)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}

				string sSQLValNum = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)this.txtNumIdCompra.Value, this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);

				if (sMensaje.ToString().Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			}
			#endregion Numero
			#region Fecha
			if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Fecha
			#region Valida identificación
			string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.cmbTipoRuc.Value);
			string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
			if (VlCed.StartsWith("Error"))
			{
				MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			#endregion Valida identificación
			#region Valida Cliente
			string sSQLValCl = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLValCl, true);
			drC.Read();
			if (!drC.HasRows)
			{
				MessageBox.Show("Cliente no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			drC.Close();
			#endregion Valida Cliente
			#region Valida Articulos en el grid
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los productos para crear la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Valida Articulos en el grid
			#region Valida Cantidad - Precio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)// || (int)dr.Cells["idArticulo"].Value > 0)
				{
					int IdArt = (int) dr.Cells["idArticulo"].Value;

					int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From ComboDetalle Where idArticulo = {0}", IdArt));

					if (iCont == 0)
					{
						decimal dPrecio = (decimal) dr.Cells["Precio"].Value;
							
						if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value)// && (int)dr.Cells["Tipo"].Value != 1) 
						{
							MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}					
					else
					{
						decimal dPrecio = (decimal) dr.Cells["Precio"].Value;
						if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value)// && (int)dr.Cells["Tipo"].Value != 1) 
						{
							MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}

					int iCantidad = (int) dr.Cells["cantidad"].Value;
					if (iCantidad == 0)
					{
						MessageBox.Show(string.Format("Digite la cantidad del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Valida Cantidad - Precio
			#region Valida Existencia
			if (bNuevo)
			{
				if (!bConsignacion)
				{
					if ((int)this.txtNumidOrigen.Value > 0 && idTipoOrigen == 43)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
						{
							if (!bExistencia((int)dr.Cells["idArticulo"].Value, true))
							{					
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}				
						}
					}
				}
			}
			#endregion Valida Existencia
			Total();
			#region Valida Datos de Entrega
			if ((bool)this.chkEntregado.Checked)
			{
				if (this.dtFechaEntrega.Value == DBNull.Value)
				{
					MessageBox.Show("Escriba la fecha de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.utcTransaccion.Tabs[3].Selected = true;
					this.dtFechaEntrega.Focus();
					return;
				}
				if (this.dtpHoraEntrega.Value == DBNull.Value)
				{
					MessageBox.Show("Escriba la hora de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.utcTransaccion.Tabs[3].Selected = true;
					this.dtpHoraEntrega.Focus();
					return;
				}
				if (this.cmbDirecciones.Value == DBNull.Value)
				{
					MessageBox.Show("Escriba la hora de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.utcTransaccion.Tabs[3].Selected = true;
					this.cmbDirecciones.Focus();
					return;
				}
			}
			#endregion Valida Datos de Entrega
			#region Verifica Total Cero
			if ((decimal) txtTotal.Value == 0.00m) 
			{
				MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Verifica Total Cero
			#region Verifica Seriales
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true))
			{
				if (!bConsignacion && (idTipoOrigen == 0 || idTipoOrigen == 43))
				{
					#region Seriales en el detalle 
					if (bSerialesL)
					{				 
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							if (dr.Cells["idArticulo"].Value != DBNull.Value)
							{
								int IdArtDet = (int) dr.Cells["idArticulo"].Value;
						
								string stMensaje = string.Format("select isNull(ManejaSerial, 0) from Articulo where idArticulo={0}", IdArtDet);
								bool bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);
								if (bManejaSerial)
								{
									if (dr.Cells["SerieL"].Value == DBNull.Value || dr.Cells["SerieL"].Value.ToString() == "")
									{
										MessageBox.Show(string.Format("Escriba la serial del artículo '{0}'.", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										this.ultraGrid1.ActiveRow = dr;
										this.ultraGrid1.ActiveRow.Selected = true;
										return;
									}
								}
							}
						}
					}
					#endregion Seriales en el detalle

					#region Seriales por grupo
					if (!bSerialesL)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							decimal VerifCantArticulo = 0;
							if (dr.Cells["idArticulo"].Value != DBNull.Value)
							{
								int IdArtDet = (int) dr.Cells["idArticulo"].Value;
								int Cant = (int) dr.Cells["Cantidad"].Value;
								string Art = (string) dr.Cells["Articulo"].Value;
								int IPosicion = (int) dr.Cells["Posicion"].Value;

								string stMensaje = string.Format("select isNull(ManejaSerial, 0) from Articulo where idArticulo={0}", IdArtDet);
								bool bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);

								if (bManejaSerial == true)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
									{
										int IdArtSerie = (int) dr1.Cells["idArticulo"].Value;
										if (IdArtSerie == IdArtDet && IPosicion == (int)dr1.Cells["Posicion"].Value)
										{
											VerifCantArticulo++;
										}
									}

									if (VerifCantArticulo < Cant)
									{
										MessageBox.Show("Ha registrado " + VerifCantArticulo + " de " + Cant + " series del artículo " + Art.Trim() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										this.ultraGrid1.ActiveRow = dr;
										this.ultraGrid1.ActiveRow.Selected = true;
										return;
									}
									else if (VerifCantArticulo > Cant)
									{
										MessageBox.Show("Ha registrado " + VerifCantArticulo + " de " + Cant + " series del artículo " + Art.Trim() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										this.ultraGrid1.ActiveRow = dr;
										this.ultraGrid1.ActiveRow.Selected = true;
										return;
									}

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
									{
										if ((int) dr1.Cells["Estado"].Value == 8 && dr1.Cells["Observaciones"].Value.ToString().Trim().Length == 0)
										{
											MessageBox.Show("Ingrese una observación para la serial " + dr1.Cells["Serial"].Value.ToString() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
											this.grdSerial.ActiveRow = dr1;
											this.grdSerial.ActiveRow.Selected = true;
											return;
										}
									}
								}
							}					
						}
					}
					#endregion Seriales por grupo
				}									
			}
			#endregion Verifica Seriales
			#region Valida Estado Origen Nota De Credito
			if ((int)this.cmbFormaPago.Value == 4)
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.txtNumidOrigen1.Value), true)) 
				{
					MessageBox.Show("La Nota De Credito Del Cliente se encuentra Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
					return;
				}

				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From BloqueoTransacciones Where idOrigen = {0} And Origen = 1 And Bloqueado = 1", (int)this.txtNumidOrigen1.Value)) > 0)
				{
					string sUsuario = "";
					string sEstacion = "";
					DateTime dtFechaBloqueo = DateTime.Today;

					string sSQLValB = string.Format("Select Usuario, Estacion, Fecha From BloqueoTransacciones Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
					SqlDataReader drBloqueo = Funcion.rEscalarSQL(cdsSeteoF, sSQLValB, true);
					drBloqueo.Read();
					if (drBloqueo.HasRows)
					{
						sUsuario = drBloqueo.GetValue(0).ToString();
						sEstacion = drBloqueo.GetValue(1).ToString();
						dtFechaBloqueo = drBloqueo.GetDateTime(2);
					}
					drBloqueo.Close();

					MessageBox.Show(string.Format("La Nota De Credito Del Cliente esta siendo Editada : \n\n Usuario : '{0}' \n Estacion : '{1}' \n Desde : '{2}' \n\n La Factura Se Cancelara hasta que el usuario termine de Editar la Nota de Credito", sUsuario, sEstacion, dtFechaBloqueo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.btnCancelar_Click(sender, e);
					return;
				}
			}
			#endregion Valida Estado Origen Nota De Credito
			#region Valida Valor Items y Cantidad de Origen
			if ((int)this.txtNumidOrigen.Value > 0)
			{
				decimal dValorOrigen = 0.00m;
				dValorOrigen = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ROUND(Total, 2) From Compra Where idCompra = {0}", (int)this.txtNumidOrigen.Value));
				if (dValorOrigen != Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2))
				{
					MessageBox.Show(string.Format("El Valor del Documento de Origen {0} no coincide con el Valor Total de la Factura {1}", dValorOrigen, Convert.ToDecimal(this.txtTotal.Value)));
					return;
				}

				int iItems = 0;
				iItems = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetCompra Where idCompra = {0}", (int)this.txtNumidOrigen.Value));
				if (iItems != this.ultraGrid1.Rows.Count)
				{
					MessageBox.Show(string.Format("El Numero de Items del Documento de Origen {0} no coincide con el Numero de Items de la Factura {1}", iItems, this.ultraGrid1.Rows.Count));
					return;
				}

				int iCantidadItems = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					iCantidadItems = iCantidadItems + (int)dr.Cells["Cantidad"].Value;
				}

				int iCantidadItemsOrigen = 0;
				iCantidadItemsOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select SUM(Cantidad) From DetCompra Where idCompra = {0}", (int)this.txtNumidOrigen.Value));
				if (iCantidadItemsOrigen != iCantidadItems)
				{
					MessageBox.Show(string.Format("La Cantidad de Productos del Documento de Origen {0} no coincide con la Cantidad de Productos de la Factura {1}", iCantidadItemsOrigen, iCantidadItems));
					return;
				}
			}
			#endregion Valida Valor Origen
			#endregion Validacion

			#region Estado Factura
			if (!bSerialesL)
			{
				if (this.lblEstado.Text.ToString() != "MANUAL") this.txtNumEstado.Value = 10;
				else this.txtNumEstado.Value = 11;
					
				int iContMan = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSer in this.grdSerial.Rows.All)
				{
					int EstManual = (int)drSer.Cells["Estado"].Value;
					if (EstManual == 8) iContMan ++;
				}
				if (iContMan > 0) this.txtNumEstado.Value = 11;
				else if (iContMan == 0) this.txtNumEstado.Value = 10;
			}
			else if (bSerialesL)
			{
				int Cont = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					Manual = (bool)dr.Cells["Manual"].Value;
					if (Manual) Cont = Cont + 1;
				}
	
				if (Cont > 0) this.txtNumEstado.Value = 11;
				else if (Cont == 0) this.txtNumEstado.Value = 10;
			}
			#endregion Estado Factura

			#region Valida Retenciones
			bool bRetencionCaja = false;
			if (BRetencion) 
			{
				this.btnRetener_Click(sender, e);
				
				if (bNuevo) bRetencionCaja = true;
				else if ((int)this.txtIdAsientoRetencion.Value == 0) bRetencionCaja = true; else bRetencionCaja = false;
			}
			else
			{					
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Retenciones Where ClienteProveedor = 0 And idCompra = {0} And Estado <> 2", (int)this.txtNumIdCompra.Value)) > 0)
				{
					if (DialogResult.Yes == MessageBox.Show("Ha Eliminado la Retención de esta Factura ¿Esta Seguro de Guardar sin Retención?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{
						string sSQLAnulaRetencion = string.Format("Exec AnulaRetencionDesdeFactura {0}", (int)this.txtNumIdCompra.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAnulaRetencion);
					}
					else return;
				}
			}
			#endregion Valida Retenciones

			bEdicion = false;

			#region Grabar
			#region Credito - CrediPOINT SIN ENTRADA - PROMOCIONALES
			if ((bCrediPoint && (decimal)this.txtCuotaDeEntrada.Value == 0.00m) || ((int)this.cmbFormaPago.Value == 6 && !bUsoAnticipo) || bFormaPagoPromocional)
			{				
				#region Valida Entrega de Articulos
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//				{
//					if ((bool)dr.Cells["Manual"].Value)
//					{
//						if ((int)Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.funArticuloSaldo({0}, {1}, GETDATE())", (int)dr.Cells["idArticulo"].Value, (int)this.cmbBodega.Value)) > 0) dr.Cells["Entregado"].Value = true;									
//						else dr.Cells["Entregado"].Value = false;
//					}
//					else dr.Cells["Entregado"].Value = true;
//				}
				#endregion Valida Entrega de Articulos

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

						if (bNuevo && bNumeracionAutomatica)
						{
							string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 0", (int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
						}
					
						#region Graba Factura
						#region Maestro
						int idDireccion = 0;
						if (this.cmbDirecciones.Text.ToString().Trim().Length > 0) idDireccion = (int)this.cmbDirecciones.Value;

						#region Fechas
						DateTime dtFecha = (DateTime)this.dtFecha.Value;
						DateTime dtFechaCad = DateTime.Today;
						DateTime dtFechaEnt = DateTime.Today;
						DateTime dtHoraEntrega = DateTime.Today;
						bool bFecCad = false;
						bool bFecEnt = false;
						bool bHorEnt = false;
						bool bFecPrimPag = false;

						if (this.cmbCaducidad.Value != System.DBNull.Value) 
						{
							bFecCad = true;
							dtFechaCad = (DateTime) this.cmbCaducidad.Value;
						}
						if ((bool)this.chkEntregado.Checked)
						{
							if (this.dtFechaEntrega.Value != System.DBNull.Value) 
							{
								bFecEnt = true;
								dtFechaCad = (DateTime) this.dtFechaEntrega.Value;
							}
							if ((string)this.dtpHoraEntrega.Value == "") 
							{
								bHorEnt = true;
								dtFechaCad = (DateTime) this.dtpHoraEntrega.Value;
							}
						}
						if (bCrediPoint) bFecPrimPag = true;
						#endregion Fechas

						sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '{54}'", 
							(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value,
							(int)this.txtIdSucursal.Value, 
							(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
							(int)this.txtNumEstado.Value,
							bConsignacion,
							(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
							(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
							(int)this.cmbVendedor.Value,
							this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
							(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
							dtFecha.ToString("yyyyMMdd"), 
							this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), 
							(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion,
							this.txtNotas.Text.ToString(),
							bCrediPoint, dTotalCuotas, NumCuotas, (decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"),
							(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
							bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado,
							dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, AutorizaEntrada, dValorCuota, (int)this.txtIdEF.Value,
							(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, this.txtNumeroCredito.Text.ToString());
						oCmdActualiza.CommandText = sSQL;
						this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Maestro
						
						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							int iPosicion = 0;
							if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicion = (int)dr.Cells["Posicion"].Value;

							string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}",
								(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value,
								(bool)dr.Cells["Manual"].Value, dr.Cells["SerieL"].Value.ToString(),
								(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
								(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
								(int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, iPosicion, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	
						
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);						
						this.ultraGrid1.DataSource = oDTDetalle;						
						#endregion Graba Factura

						int idDetFormaPago = 0;
						DateTime dtFecCaducidad = dtFecVencCupon;// DateTime.Today;
				
						#region Retenciones
						if (BRetencion)
						{
							#region Crea Retencion
							DateTime dtFechaRet = (DateTime) this.dtFechaRetencion.Value;
											
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
							{
								if (!(bool) dr.Cells["IVARenta"].Value) dRetenidoIVA = dRetenidoIVA + (decimal)dr.Cells["Valor"].Value;
								if ((bool) dr.Cells["IVARenta"].Value) dRetenidoRenta = dRetenidoRenta + (decimal)dr.Cells["Valor"].Value;				
								dRetenido = dRetenido + (decimal)dr.Cells["Valor"].Value;
							}
							
							#region Guarda Encabezado Retencion
							string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 0, 0, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
								(int) this.txtidRetencion.Value, (int)this.txtNumIdCompra.Value, this.txtSerieRetencion.Text.ToString(), this.txtNumeroRetencion.Text.ToString(), this.txtAutorizacionRetencion.Text.ToString(), 
								dtFechaCadR.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), "", dRetenidoIVA, dRetenidoRenta, dRetenido);
							oCmdActualiza.CommandText = sSQLGR;
							this.txtidRetencion.Value = (int)oCmdActualiza.ExecuteScalar();
							#endregion Guarda Encabezado Retencion

							#region Valida Si existe Detalle de Retencion
							if (!bNuevo)
							{
								string sSQLValida = string.Format("Exec ValidaDetalleDeRetencion {0}", (int)this.txtidRetencion.Value);
								oCmdActualiza.CommandText = sSQLValida;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Valida Si existe Detalle de Retencion

							#region Guarda Detalle Retencion
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
							{
								string sSQLGDR = string.Format("Exec GuardaDetalleRetenciones {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
									(int)dr.Cells["idDetalleRetencion"].Value, (int)this.txtidRetencion.Value, (bool)dr.Cells["IVARenta"].Value, (int)dr.Cells["idRetencionIVARenta"].Value, (decimal)dr.Cells["Base"].Value, 
									(decimal)dr.Cells["Valor"].Value, (decimal)dr.Cells["Porcentaje"].Value);
								oCmdActualiza.CommandText = sSQLGDR;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Guarda Detalle Retencion

							#region Actualiza grid de Retenciones
							string sSQLActDetalleRet = string.Format("Exec ConsultaRetDetalle {0}", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLActDetalleRet;
							SqlDataAdapter oDADetalleRet = new SqlDataAdapter(oCmdActualiza);
							DataTable oDTDetalleRet = new DataTable();
							oDADetalleRet.Fill(oDTDetalleRet);
							this.ultraGrid2.DataSource = oDTDetalleRet;							
							#endregion Actualiza grid de Retenciones

							string sSQLRetenido = string.Format("Update Compra Set Retenido = {0}, RetenidoIVA = {2}, RetenidoRenta = {3} Where idCompra = {1}", dRetenido, (int)this.txtNumIdCompra.Value, dRetenidoIVA, dRetenidoRenta);
							oCmdActualiza.CommandText = sSQLRetenido;
							oCmdActualiza.ExecuteNonQuery();

							#region Guarda retencion en caja
							if (!bNuevo)
							{
								string sSQLValRetExist = string.Format("Exec ValidaExistenciaDFP {0}, 10", (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLValRetExist;
								idDetFormaPagoRetencion = (int)oCmdActualiza.ExecuteScalar();
								
								if (idDetFormaPagoRetencion > 0)
								{
									#region Actualiza en Caja el Registro de la Retencion 
									string sSQLCobroRet = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
										idDetFormaPagoRetencion, (int)this.txtNumIdCompra.Value, 1, 10, dRetenido, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
									oCmdActualiza.CommandText = sSQLCobroRet;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza en Caja el Registro de la Retencion
								}
								else
								{
									#region Valida si la Retencion se creo en Modulo de Facturacion o en Cuentas Por Cobrar									
									if ((int)this.txtIdAsientoRetencion.Value == 0)
									{
										string sSQLCobroRet = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
											idDetFormaPagoRetencion, (int)this.txtNumIdCompra.Value, 1, 10, dRetenido, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
										oCmdActualiza.CommandText = sSQLCobroRet;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Valida si la Retencion se creo en Modulo de Facturacion o en Cuentas Por Cobrar
								}
							}
							else							
							{
								#region Si es nuevo el registro de Factura Guarda en caja el registro de la retencion
								string sSQLCobroRet = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
									idDetFormaPagoRetencion, (int)this.txtNumIdCompra.Value, 1, 10, dRetenido, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
								oCmdActualiza.CommandText = sSQLCobroRet;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Si es nuevo el registro de Factura Guarda en caja el registro de la retencion
							}
							#endregion Guarda retencion en caja

							#region Asiento de Retencion
							if ((int)this.txtIdAsientoRetencion.Value > 0 && !bNuevo && !bRetencionCaja)
							{																		
								string sSQLAsientoR = string.Format("Exec CreaAsientoRetencionVenta {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, '{10}'", 
									this.txtIdAsientoRetencion.Value, (int)this.txtidRetencion.Value, (int)this.txtNumIdCliente.Value, (int)this.cmbFormaPago.Value, this.txtNombre.Text.ToString(), this.txtNumeroRetencion.Text.ToString(), 
									this.txtNumero.Text.ToString(), dtFechaRet.ToString("yyyyMMdd"), (int)this.txtNumIdProyecto.Value, dRetenido, this.txtRuc.Text.ToString());
								oCmdActualiza.CommandText = sSQLAsientoR;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Asiento de retencion
							#endregion Crea Retencion
						}
						else
						{
							if (!bNuevo) 
							{
								string sSQLActualiza = string.Format("Update Compra Set Retenido = 0, RetenidoIVA = 0, RetenidoRenta = 0 Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLActualiza;
								oCmdActualiza.ExecuteNonQuery();									
							}
						}
						#endregion Retenciones
						
						#region Guarda Cobro
						if (!bNuevo)
						{
							string sSQLIdDetFormaPago = string.Format("Exec ValidaExistenciaDFP {0}, {1}", (int)this.txtNumIdCompra.Value, (int)this.cmbFormaPago.Value);
							oCmdActualiza.CommandText = sSQLIdDetFormaPago;
							idDetFormaPago = (int)oCmdActualiza.ExecuteScalar(); 
						}

						decimal dTotalCobro = Convert.ToDecimal(this.txtTotal.Value);
						if (bRetencionCaja) dTotalCobro = Convert.ToDecimal(this.txtTotal.Value) - dRetenido;

						if (!bNuevo && BRetencion)
						{
							string sSQLIdDetFormaPago = string.Format("Exec ValidaExistenciaDFP {0}, 10", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLIdDetFormaPago;
							int ValIdDetFormaPago = (int)oCmdActualiza.ExecuteScalar(); 							
						}

						string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
							idDetFormaPago, (int)this.txtNumIdCompra.Value, 1, (int)this.cmbFormaPago.Value, dTotalCobro, 0, 0, "", sNumero, dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
						oCmdActualiza.CommandText = sSQLCobro;
						oCmdActualiza.ExecuteNonQuery();						
						#endregion Guarda Cobro
						
						#region Saldo Factura
						string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", (int)this.txtNumIdCompra.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value);
						oCmdActualiza.CommandText = sSQLSaldo;
						oCmdActualiza.ExecuteNonQuery();
						#endregion Saldo Factura
						
						#region Anticipos
						if (bUsoAnticipo)
						{	
							#region Actualiza saldo anticipos
							if ((decimal)this.txtTotal.Value == dAnticipo)
							{
								sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
							
								MessageBox.Show("Factura Cancelada con anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							else if ((decimal)this.txtTotal.Value < dAnticipo)
							{
								dAnticipo = (decimal)this.txtTotal.Value;
								sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();

								MessageBox.Show("Factura Cancelada con anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							else if ((decimal)this.txtTotal.Value > dAnticipo)
							{
								sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Actualiza saldo anticipos
						}
						#endregion Anticipos
						
						#region Seriales
//						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true))
//						{						
							if (idTipoOrigen == 13 || idTipoOrigen == 11)
							{
								if (bNuevo)
								{
									#region Seriales
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
									{
										string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
										oCmdActualiza.CommandText = sSQLMovSerial;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Seriales

									string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQLActSeriales;
									SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
									DataTable oDTSeriales = new DataTable();
									oDASeriales.Fill(oDTSeriales);

									this.grdSerial.DataSource = oDTSeriales;

									string sSQLConteoActivo = string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(Date, Fecha) <= '{1}' And Estado = 1", (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"));
									oCmdActualiza.CommandText = sSQLConteoActivo;
									int iConteoActivo = (int)oCmdActualiza.ExecuteScalar();

									if (iConteoActivo > 0)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
										{
											string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
											oCmdActualiza.CommandText = sSQLConteo;
											oCmdActualiza.ExecuteNonQuery();
										}
									}
								}
							}
							else
							{
								if (!bConsignacion) // || bCrediPoint)
								{									
									#region Grupos de seriales
									if (!bSerialesL)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
										{
											if (dr.Cells["idArticulo"].Value != DBNull.Value && (int) dr.Cells["idArticulo"].Value > 0)
											{
												int IdArticulo = (int) dr.Cells["idArticulo"].Value;
												int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;										
												int IPosicion = (int) dr.Cells["Posicion"].Value;
						
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
												{
													int IdArtSerial = (int) dr1.Cells["idArticulo"].Value;
													int IEstado = (int) dr1.Cells["Estado"].Value;
											
													if (IdArticulo == IdArtSerial && IPosicion == (int) dr1.Cells["Posicion"].Value)
													{
														if (IEstado == 8)
														{
															int idSerial = 0;
															if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;
 
															string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 8, {3}", 
																idSerial, (int)dr1.Cells["idArticulo"].Value, dr1.Cells["Serial"].Value.ToString(), (int)this.cmbBodega.Value);
															oCmdActualiza.CommandText = sSQLGSeriales;
															int NIdSerial = (int)oCmdActualiza.ExecuteScalar();

															string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'LA SERIAL NO TIENE REGISTRO EN EL INVENTARIO.', 12, 0, {4}, {5}", 
																NIdSerial, IdDetCompra, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int) dr1.Cells["Posicion"].Value);
															oCmdActualiza.CommandText = sSQLGDetalleSeriales;
															oCmdActualiza.ExecuteNonQuery();
														}
														else if (IEstado == 4 || IEstado == 2)
														{
															int idSerial = 0;
															int idDetSerial = 0;

															if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
															if (dr1.Cells["idDetalleSerial"].Value != DBNull.Value) idDetSerial = (int)dr1.Cells["idDetalleSerial"].Value;
														
															string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", 
																idSerial, IdDetCompra, idDetSerial, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int) dr1.Cells["Posicion"].Value);
															oCmdActualiza.CommandText = sSQLGDetalleSeriales;
															oCmdActualiza.ExecuteNonQuery();
															
															string sSQLActEstSerial = string.Format("Update Seriales Set Estado = {0} Where idSerial = {1}", 4, idSerial);
															oCmdActualiza.CommandText = sSQLActEstSerial;															
															oCmdActualiza.ExecuteNonQuery();
																														
															string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
																dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
															oCmdActualiza.CommandText = sSQLConteo;
															oCmdActualiza.ExecuteNonQuery();
														}
													}
												}
											}
										}

										string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
										oCmdActualiza.CommandText = sSQLActSeriales;
										SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
										DataTable oDTSeriales = new DataTable();
										oDASeriales.Fill(oDTSeriales);

										this.grdSerial.DataSource = oDTSeriales;										
									}
									#endregion Grupos de seriales
								}
							}
//						}						
						#endregion Seriales

						#region Estado Documento de origen
						if ((int)this.txtNumidOrigen.Value > 0)
						{
							if (idTipoOrigen == 43)
							{
								string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado = 5", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
								oCmdActualiza.CommandText = sSQLEstado;
								oCmdActualiza.ExecuteNonQuery();
							}
							else if (idTipoOrigen == 11 || idTipoOrigen == 13)
							{
								string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado In (3, 5)", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
								oCmdActualiza.CommandText = sSQLEstado;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
						#endregion Estado Documento de origen
						
						#region Pagos - Codigos de barras
						if (idTipoOrigen != 43)
						{
							#region Genera pagos desde pedido
							sSQL = string.Format("Exec GeneraPlanDePagosCredito {0}", (int)this.txtNumIdCompra.Value); //GeneraPagos
							oCmdActualiza.CommandText = sSQL;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Genera pagos desde pedido
											
							#region Codigo de barras
							string sSQLCB = "";
							string sCodigoBarras = "";
							int iCont=0;
							string sSQLValCB = string.Format("Select CodigoBarras From Cliente Where idCliente = {0} And Proveedor = 0", (int)this.txtNumIdCliente.Value);
							oCmdActualiza.CommandText = sSQLValCB;
							bool bCodigoBarras = (bool)oCmdActualiza.ExecuteScalar();
							
							if (bCodigoBarras)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
								{
									sSQLCB = string.Format("Exec RetornaCodigoDeBarras {0}, {1}", (int)this.txtNumIdCliente.Value, (int)dr.Cells["idArticulo"].Value);
									oCmdActualiza.CommandText = sSQLCB;
									sCodigoBarras = oCmdActualiza.ExecuteScalar().ToString();// Funcion.sEscalarSQL(cdsSeteoF, sSQLCB);

									if (sCodigoBarras.Length > 0)
									{
										string sSQLActCB = string.Format("Update DetCompra Set CodigoBarras = '{0}' Where idArticulo = {1} And idDetCompra = {2}", 
											sCodigoBarras, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetCompra"].Value);
										oCmdActualiza.CommandText = sSQLActCB;
										oCmdActualiza.ExecuteNonQuery();
									}								
								}
							}
							#endregion Codigo de barras
						}
						#endregion Pagos - Codigos de barras

						#region Crea Nota de entrega ZOOM
						if (Funcion.bEscalarSQL(cdsSeteoF, "Select CreaNotaEntrega From SeteoF", true) && bNuevo)
						{
							string sSQLNE = string.Format("Exec CreaNotaDeEntrega {0}", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLNE;
							int idNE = (int)oCmdActualiza.ExecuteScalar();

							string sSQLDNE = string.Format("Exec CreaDetalleNotaDeEntrega {0}, {1}", (int)this.txtNumIdCompra.Value, idNE);
							oCmdActualiza.CommandText = sSQLDNE;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Crea Nota de entrega ZOOM
						
						#region Asiento Automatico
						if (bAsientoAutomatico) 
						{						
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", this.txtNumIdCompra.Value.ToString());
							oCmdActualiza.CommandText = sSQLAsiento;
							oCmdActualiza.ExecuteNonQuery(); 

							string sSQLRetIdAsiento = string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLRetIdAsiento; 
							this.txtNumIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar(); 
						}
						#endregion Asiento Automatico

						if (idBloqueaTransacciones > 0)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						oTransaction.Commit();

						Cursor = Cursors.Default;							

						#region Impresión
						if (bNuevo)
						{
							using (frmMensajeNumeración miMsje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, 0, false))
							{
								if (DialogResult.OK == miMsje.ShowDialog())
								{
									if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVIG'")) Imprimir(1);
								}
							}
						}
						#endregion Impresión
	
						#region Controles
						if ((int)this.txtNumEstado.Value == 11)	this.lblEstado.Text = "MANUAL";
						else if ((int)this.txtNumEstado.Value == 10) this.lblEstado.Text = "CONTABILIZADO";

						this.cmbFormaPago.Enabled = false;
						this.cmbBodega.Enabled = false;
						this.cmbVendedor.Enabled = false;
						this.cmbTipoRuc.Enabled = false;

						this.txtNumero.Enabled = false;
						this.txtRuc.Enabled = false;
						this.dtFecha.Enabled = false;
						this.txtNumeroCredito.Enabled = false;

						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (!miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
						if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						this.btnGuardar.Enabled = false;
						this.btnCancelar.Enabled = true;
						this.btnImprimeDctos.Enabled = false;
						this.btnClaveDescuento.Enabled = false;
						this.btnTemporal.Enabled = false;
						this.btnCliente.Enabled = false;

						this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieL"].CellActivation = Activation.AllowEdit;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.AllowEdit;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;
					
						this.ModificaGrilla(false);
						this.ModificaGrillaSerial(false);
						HabilitaTab(false);
						this.grdSerial.Enabled = true;
						bNuevo = false;
						bEdicion = false;
						bDescuentoActivo = false;
						idBloqueaTransacciones = 0;
						#endregion Controles
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}

				return;
			}
			#endregion Credito - CrediPOINT SIN ENTRADA - PROMOCIONALES

			#region Efectivo - Tarjeta 
			if (bCrediPoint) dValorCrediPoint = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
			using (VentaFormaDePago miCobro = new VentaFormaDePago ((int)this.txtNumIdCompra.Value, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, this.txtComprobante.Text.ToString(), BUsoNotaDeCredito, (int)this.txtNumidOrigen1.Value, dSaldoNotaDeCredito, bRetencionCaja, dValorRetencion, bUsoAnticipo, dAnticipo, bCrediPoint, dValorCrediPoint, (decimal)this.txtCuotaDeEntrada.Value, 1, (int)this.txtNumIdCliente.Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, (int)this.cmbBodega.Value, this.ultraGrid1, (DateTime)this.dtFecha.Value, false))
			{
				if (DialogResult.OK == miCobro.ShowDialog())
				{
					#region Valida Entrega de Articulos
//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//					{
//						if ((bool)dr.Cells["Manual"].Value)
//						{
//							if ((int)Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.funArticuloSaldo({0}, {1}, GETDATE())", (int)dr.Cells["idArticulo"].Value, (int)this.cmbBodega.Value)) > 0) dr.Cells["Entregado"].Value = true;									
//							else dr.Cells["Entregado"].Value = false;
//						}
//						else dr.Cells["Entregado"].Value = true;
//					}
					#endregion Valida Entrega de Articulos
						
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

							if (bNuevo && bNumeracionAutomatica)
							{
								string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 0", (int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLNumero;
								this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
							}	
			
							#region Graba Factura
							#region Maestro
							int idDireccion = 0;
							if (this.cmbDirecciones.Text.ToString().Trim().Length > 0) idDireccion = (int)this.cmbDirecciones.Value;

							#region Fechas
							DateTime dtFecha = (DateTime)this.dtFecha.Value;
							DateTime dtFechaCad = DateTime.Today;
							DateTime dtFechaEnt = DateTime.Today;
							DateTime dtHoraEntrega = DateTime.Today;
							bool bFecCad = false;
							bool bFecEnt = false;
							bool bHorEnt = false;
							bool bFecPrimPag = false;
							if (this.cmbCaducidad.Value != System.DBNull.Value) 
							{
								bFecCad = true;
								dtFechaCad = (DateTime) this.cmbCaducidad.Value;
							}
							if (this.dtFechaEntrega.Value != System.DBNull.Value) 
							{
								bFecEnt = true;
								dtFechaCad = (DateTime) this.dtFechaEntrega.Value;
							}
							if ((string)this.dtpHoraEntrega.Value == "") 
							{
								bHorEnt = true;
								//								dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
							}
							if (bCrediPoint) bFecPrimPag = true;
							#endregion Fechas

							sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '{54}'", 
								(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value,
								(int)this.txtIdSucursal.Value, 
								(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
								(int)this.txtNumEstado.Value,
								bConsignacion,
								(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
								(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
								(int)this.cmbVendedor.Value,
								this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
								(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
								dtFecha.ToString("yyyyMMdd"), 
								this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), 
								(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion,
								this.txtNotas.Text.ToString(),
								bCrediPoint, dTotalCuotas, NumCuotas, (decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"),
								(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
								bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado, 
								dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, AutorizaEntrada, dValorCuota, (int)this.txtIdEF.Value,
								(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, this.txtNumeroCredito.Text.ToString());
							oCmdActualiza.CommandText = sSQL;
							this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
							#endregion Maestro

							#region Detalle
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{	
								int iPosicion = 0;
								if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicion = (int)dr.Cells["Posicion"].Value;

								string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}",
									(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value,
									(bool)dr.Cells["Manual"].Value, dr.Cells["SerieL"].Value.ToString(),
									(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
									(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
									(int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, iPosicion, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value);
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Detalle	

							string sSQLActDetalle = string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLActDetalle;
							SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
							DataTable oDTDetalle = new DataTable();
							oDADetalle.Fill(oDTDetalle);
							this.ultraGrid1.DataSource = oDTDetalle;							
							#endregion Graba Factura
					
							#region Retenciones
							if (BRetencion)
							{
								#region Crea Retencion
								DateTime dtFechaRet = (DateTime) this.dtFechaRetencion.Value;
									
								#region Valores de Retencion
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
								{
									if (!(bool) dr.Cells["IVARenta"].Value) dRetenidoIVA = dRetenidoIVA + Convert.ToDecimal(dr.Cells["Valor"].Value);
									if ((bool) dr.Cells["IVARenta"].Value) dRetenidoRenta = dRetenidoRenta + Convert.ToDecimal(dr.Cells["Valor"].Value);				
									dRetenido = dRetenido + Convert.ToDecimal(dr.Cells["Valor"].Value);
								}
								#endregion Valores de Retencion	

								#region Guarda Encabezado Retencion
								string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 0, 0, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
									(int) this.txtidRetencion.Value, (int)this.txtNumIdCompra.Value, this.txtSerieRetencion.Text.ToString(), this.txtNumeroRetencion.Text.ToString(), this.txtAutorizacionRetencion.Text.ToString(), 
									dtFechaCadR.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), "", dRetenidoIVA, dRetenidoRenta, dRetenido);
								oCmdActualiza.CommandText = sSQLGR;
								this.txtidRetencion.Value = (int)oCmdActualiza.ExecuteScalar();
								#endregion Guarda Encabezado Retencion

								#region Valida Si existe Detalle de Retencion
								if (!bNuevo)
								{
									string sSQLValida = string.Format("Exec ValidaDetalleDeRetencion {0}", (int)this.txtidRetencion.Value);
									oCmdActualiza.CommandText = sSQLValida;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Valida Si existe Detalle de Retencion

								#region Guarda Detalle Retencion
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
								{
									string sSQLGDR = string.Format("Exec GuardaDetalleRetenciones {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
										(int)dr.Cells["idDetalleRetencion"].Value, (int)this.txtidRetencion.Value, (bool)dr.Cells["IVARenta"].Value, (int)dr.Cells["idRetencionIVARenta"].Value, 
										(decimal)dr.Cells["Base"].Value, (decimal)dr.Cells["Valor"].Value, (decimal)dr.Cells["Porcentaje"].Value);
									oCmdActualiza.CommandText = sSQLGDR;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Guarda Detalle Retencion

								#region Actualiza grid de Retenciones
								string sSQLActDetalleRet = string.Format("Exec ConsultaRetDetalle {0}", (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLActDetalleRet;
								SqlDataAdapter oDADetalleRet = new SqlDataAdapter(oCmdActualiza);
								DataTable oDTDetalleRet = new DataTable();
								oDADetalleRet.Fill(oDTDetalleRet);
								this.ultraGrid2.DataSource = oDTDetalleRet;							
								#endregion Actualiza grid de Retenciones

								string sSQLRetenido = string.Format("Update Compra Set Retenido = {0}, RetenidoIVA = {2}, RetenidoRenta = {3} Where idCompra = {1}", dRetenido, (int)this.txtNumIdCompra.Value, dRetenidoIVA, dRetenidoRenta);
								oCmdActualiza.CommandText = sSQLRetenido;
								oCmdActualiza.ExecuteNonQuery();

								#region Asiento de Retencion solo al Modificar la Factura
								if ((int)this.txtIdAsientoRetencion.Value > 0 && !bNuevo && !bRetencionCaja)
								{																		
									string sSQLAsientoR = string.Format("Exec CreaAsientoRetencionVenta {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, '{10}'", 
										this.txtIdAsientoRetencion.Value, (int)this.txtidRetencion.Value, (int)this.txtNumIdCliente.Value, (int)this.cmbFormaPago.Value, this.txtNombre.Text.ToString(), this.txtNumeroRetencion.Text.ToString(), 
										this.txtNumero.Text.ToString(), dtFechaRet.ToString("yyyyMMdd"), (int)this.txtNumIdProyecto.Value, dRetenido, this.txtRuc.Text.ToString());
									oCmdActualiza.CommandText = sSQLAsientoR;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Asiento de retencion solo al Modificar la Factura
								#endregion Crea Retencion
							}
							else
							{
								#region Actualiza Valores de Retencion
								if (!bNuevo) 
								{
									string sSQLActualiza = string.Format("Update Compra Set Retenido = 0, RetenidoIVA = 0, RetenidoRenta = 0 Where idCompra = {0} ", (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQLActualiza;
									oCmdActualiza.ExecuteNonQuery();

									string sSQLValRetExist = string.Format("Exec ValidaExistenciaDFP {0}, 10", (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQLValRetExist;
									idDetFormaPagoRetencion = (int)oCmdActualiza.ExecuteScalar();
								
									if (idDetFormaPagoRetencion > 0)
									{
										string sSQLEliRetCaja = string.Format("Delete From DetalleFormaPago Where idDetFormaPago = {0} ", idDetFormaPagoRetencion);
										oCmdActualiza.CommandText = sSQLEliRetCaja;
										oCmdActualiza.ExecuteNonQuery();
									}
								}
								#endregion Actualiza Valores de Retencion
							}
							#endregion Retenciones

							#region Guarda Cobro de la factura	
							if (miCobro.ultraGrid1.Rows.Count > 0)
							{
								dAnticipo = 0.00m;								
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
								{
									int idDetFormaPago = 0;
									int idBanco = 0;
									int idTarjeta = 0;
									int idCuenta = 0;
									int idPlan = 0;
									int idPlazo = 0;
									decimal dComTarjeta = 0.00m;
									int idDocumento = 0;
									DateTime dtFecCaducidad = DateTime.Today;
							
									if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
									if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
									if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
									if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
									if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
									if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
									if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
									if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
								
									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}", 
										idDetFormaPago, 
										(int)this.txtNumIdCompra.Value, 1, 
										(int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
										idBanco, idTarjeta,
										dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),								
										dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, 
										dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
										idPlan, idPlazo, dComTarjeta, idDocumento);
									oCmdActualiza.CommandText = sSQLCobro;
									oCmdActualiza.ExecuteNonQuery();									

									/* VALOR DE ANTICIPOS */
									if ((int)dr.Cells["idFormaPago"].Value == 8) dAnticipo = dAnticipo + Convert.ToDecimal(dr.Cells["Valor"].Value);

									/* ACTUALIZA SALDO DE NOTA DE CREDITO */
									if (idDocumento > 0)
									{
										oCmdActualiza.CommandText = string.Format("Exec ActualizaSaldoFacturas {0}, 5, 4", idDocumento, (int)this.txtIdTipoFactura.Value, 4);
										oCmdActualiza.ExecuteNonQuery();										
									}
								}								
							}
							#endregion Guarda Cobro de la factura

							#region Saldo Factura
							if ((int)this.cmbFormaPago.Value == 9)
							{
								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", (int)this.txtNumIdCompra.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbFormaPago.Value);
								oCmdActualiza.CommandText = sSQLSaldo;
								oCmdActualiza.ExecuteNonQuery();								
							}
							#endregion Saldo Factura

							#region Anticipos
							if (bUsoAnticipo)
							{	
								#region Actualiza saldo anticipos
								if ((decimal)this.txtTotal.Value == dAnticipo)
								{
									sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery();
							
									MessageBox.Show("Factura Cancelada con anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
								else if ((decimal)this.txtTotal.Value < dAnticipo)
								{
									dAnticipo = (decimal)this.txtTotal.Value;
									sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery();

									MessageBox.Show("Factura Cancelada con anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
								else if ((decimal)this.txtTotal.Value > dAnticipo)
								{
									sSQL = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)this.txtNumIdCliente.Value, dAnticipo, (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Actualiza saldo anticipos
							}
							#endregion Anticipos

							#region Seriales
//							if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true))
//							{
								if (idTipoOrigen == 13 || idTipoOrigen == 11)
								{
									if (bNuevo)
									{
										#region Seriales
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
										{
											string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
											oCmdActualiza.CommandText = sSQLMovSerial;
											oCmdActualiza.ExecuteNonQuery();										
										}
										#endregion Seriales

										string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
										oCmdActualiza.CommandText = sSQLActSeriales;
										SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
										DataTable oDTSeriales = new DataTable();
										oDASeriales.Fill(oDTSeriales);		
										this.grdSerial.DataSource = oDTSeriales;

										string sSQLConteoActivo = string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(Date, Fecha) <= '{1}' And Estado = 1", (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"));
										oCmdActualiza.CommandText = sSQLConteoActivo;
										int iConteoActivo = (int)oCmdActualiza.ExecuteScalar();

										if (iConteoActivo > 0)
										{
											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
											{
												string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
												oCmdActualiza.CommandText = sSQLConteo;
												oCmdActualiza.ExecuteNonQuery();
											}
										}
									}
								}
								else
								{
									if (!bConsignacion)
									{
										#region Grupos de seriales
										if (!bSerialesL)
										{
											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
											{
												if (dr.Cells["idArticulo"].Value != DBNull.Value && (int) dr.Cells["idArticulo"].Value > 0)
												{
													int IdArticulo = (int) dr.Cells["idArticulo"].Value;
													int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;										
													int IPosicion = (int) dr.Cells["Posicion"].Value; 					

													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
													{
														int IdArtSerial = (int) dr1.Cells["idArticulo"].Value;
														int IEstado = (int) dr1.Cells["Estado"].Value;
											
														if (IdArticulo == IdArtSerial && IPosicion == (int) dr1.Cells["Posicion"].Value)
														{
															if (IEstado == 12)
															{
																int idSerial = 0;
																int NIdSerial = 0;
																if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;
																
																if (idSerial < 1)
																{
																	string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 12, {3}", idSerial, (int)dr1.Cells["idArticulo"].Value, dr1.Cells["Serial"].Value.ToString(), (int)this.cmbBodega.Value);
																	oCmdActualiza.CommandText = sSQLGSeriales;
																	NIdSerial = (int)oCmdActualiza.ExecuteScalar();
																}
																else NIdSerial = idSerial;

																string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'LA SERIAL NO TIENE REGISTRO EN EL INVENTARIO.', 12, 0, {4}, {5}", NIdSerial, IdDetCompra, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
																oCmdActualiza.CommandText = sSQLGDetalleSeriales;
																oCmdActualiza.ExecuteNonQuery();

																oCmdActualiza.CommandText = string.Format("Update Seriales Set Estado = {0} Where idSerial = {1}", 12, idSerial);
																oCmdActualiza.ExecuteNonQuery();

																string sSQLSM = string.Format("Insert Into SerialesManuales (idSerial, Serial, idArticulo, Bodega) Values ({0}, '{1}', {2}, {3})", 
																	NIdSerial, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value, (int)this.cmbBodega.Value);
																oCmdActualiza.CommandText = sSQLSM;
																oCmdActualiza.ExecuteNonQuery();	
															}
															else if (IEstado == 4 || IEstado == 2)
															{
																int idSerial = 0;
																int idOrigen = 0;

																if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
																if (dr1.Cells["idOrigen"].Value != DBNull.Value) idOrigen = (int)dr1.Cells["idOrigen"].Value;

																string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", idSerial, IdDetCompra, idOrigen, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
																oCmdActualiza.CommandText = sSQLGDetalleSeriales;
																oCmdActualiza.ExecuteNonQuery();
																
																string sSQLActEstSerial = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", 4, idSerial, (int)this.cmbBodega.Value);
																oCmdActualiza.CommandText = sSQLActEstSerial;																
																oCmdActualiza.ExecuteNonQuery();

																string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
																oCmdActualiza.CommandText = sSQLConteo;
																oCmdActualiza.ExecuteNonQuery();
															}
														}
													}
												}
											}

											string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLActSeriales;
											SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
											DataTable oDTSeriales = new DataTable();
											oDASeriales.Fill(oDTSeriales);

											this.grdSerial.DataSource = oDTSeriales;											
										}
										#endregion Grupos de seriales
									}
								}
							//}
							#endregion Seriales

							#region Estado Documento de origen
							if ((int)this.txtNumidOrigen.Value > 0)
							{
								if (idTipoOrigen == 43)
								{
									string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado = 5", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
									oCmdActualiza.CommandText = sSQLEstado;
									oCmdActualiza.ExecuteNonQuery();
								}
								else if (idTipoOrigen == 11 || idTipoOrigen == 13)
								{
									string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado In (3, 5)", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
									oCmdActualiza.CommandText = sSQLEstado;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							#endregion Estado Documento de origen

							#region Crea Nota de entrega ZOOM
							if (Funcion.bEscalarSQL(cdsSeteoF, "Select CreaNotaEntrega From SeteoF", true) && bNuevo)
							{
								string sSQLNE = string.Format("Exec CreaNotaDeEntrega {0}", (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLNE;
								int idNE = (int)oCmdActualiza.ExecuteScalar();

								string sSQLDNE = string.Format("Exec CreaDetalleNotaDeEntrega {0}, {1}", (int)this.txtNumIdCompra.Value, idNE);								
								oCmdActualiza.CommandText = sSQLDNE;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Crea Nota de entrega ZOOM

							#region Asiento Automatico
							if (bAsientoAutomatico) 
							{						
								string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", this.txtNumIdCompra.Value.ToString());
								oCmdActualiza.CommandText = sSQLAsiento;
								oCmdActualiza.ExecuteNonQuery();

								string sSQLRetIdAsiento = string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLRetIdAsiento; 
								this.txtNumIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar(); 
								
								Cursor = Cursors.Default;								
							}
							#endregion Asiento Automatico

							if (idBloqueaTransacciones > 0)
							{
								string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLBloqueo;
								oCmdActualiza.ExecuteNonQuery();
							}

							oTransaction.Commit();	
					
							#region Impresión
							if (bNuevo)
							{
								using (frmMensajeNumeración miMsje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, 0, false))
								{
									if (DialogResult.OK == miMsje.ShowDialog())
									{
										if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVIG'")) Imprimir(1);
									}
								}
							}
							#endregion Impresión

							if ((int)this.txtNumEstado.Value == 11)	this.lblEstado.Text = "MANUAL";
							else if ((int)this.txtNumEstado.Value == 10) this.lblEstado.Text = "CONTABILIZADO";
					
							#region Controles
							this.cmbFormaPago.Enabled = false;
							this.cmbBodega.Enabled = false;
							this.cmbVendedor.Enabled = false;
							this.cmbTipoRuc.Enabled = false;

							this.txtNumero.Enabled = false;
							this.txtRuc.Enabled = false;
							this.dtFecha.Enabled = false;
							this.txtNumeroCredito.Enabled = false;

							if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
							if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
							if (miAcceso.Editar) this.btnEditar.Enabled = true;
							if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
							else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							{
								if ((int)this.cmbFormaPago.Value != 9)
								{
									if ((DateTime)this.dtFecha.Value == DateTime.Today) this.btnAnular.Enabled = true;
								}
							}
							if (!miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							this.btnGuardar.Enabled = false;
							this.btnCancelar.Enabled = true;
							this.btnImprimeDctos.Enabled = false;
							this.btnClaveDescuento.Enabled = false;
							this.btnTemporal.Enabled = false;
              this.btnCliente.Enabled = false;

							this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieL"].CellActivation = Activation.AllowEdit;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.AllowEdit;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;
					
							ModificaGrilla(false);
							this.ModificaGrillaSerial(false);
							HabilitaTab(false);
							this.grdSerial.Enabled = true;
							bNuevo = false;
							bEdicion = false;
							bDescuentoActivo = false;
							idBloqueaTransacciones = 0;
							#endregion Controles
						}
						catch (Exception ex)
						{
							try
							{
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex2)
							{
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							}
						}
						finally
						{
							oConexion.Close();
						}
					}
				}
				else bEdicion = true;
			}
			#endregion Efectivo - Tarjeta 
			#endregion Grabar
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			this.Consultar((int)this.txtNumIdCompra.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede editar FACTURA ANULADA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
			}
			else if (!miAcceso.Editar)
			{								
				MessageBox.Show("No tiene Acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else if ((int)this.txtNumEstado.Value != 6 && !miAcceso.EditarContabilizado)
			{								
				MessageBox.Show("No puede editar una Factura Contabilizada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}			
			else if (!miAcceso.EditaImpresa && bImpreso)
			{
				MessageBox.Show("No puede Editar Facturas Impresas.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("No puede Editar la Factura esta Asociada a una Retencion que a Afecta a Multiples Facturas, \n\nAnule la Retencion o Consulte al Administrador", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Bloquear, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true))
			{
				MessageBox.Show("No puede Editar la Factura esta Bloqueada", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}	
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Editar esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);												
			}
			else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
				DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
				if ((miAcceso.EditarSoloDia && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Estado, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) != 6) && dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
				{
					MessageBox.Show("Solo puede editar Facturas de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				}
				else
				{
					string sNumeroNC = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VentasValidaNotaDeCredito {0}", (int)this.txtNumIdCompra.Value));
					if (sNumeroNC.Length > 0)
					{
						MessageBox.Show(string.Format("Esta Factura esta Asociada con la Nota de Credito N. '{0}'", sNumeroNC), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}						
					else if ((int)this.cmbFormaPago.Value == 6 && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec VentaValidaPagos {0}", (int)this.txtNumIdCompra.Value)) > 0)
					{
						MessageBox.Show("Esta Factura ya tiene Pagos Registrados ", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						if ((int)this.txtNumidOrigen.Value > 0) idTipoOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idTipoFactura, 0) From Compra Where idCompra = {0}", this.txtNumidOrigen.Value));

						if (miAcceso.BEditarNombre)
						{
							if (!this.bCrediPoint)
							{
								this.cmbTipoRuc.Enabled = true;
								this.txtRuc.Enabled = true;
								this.btnCliente.Enabled = true;
							}
						}

						if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
						if (miAcceso.CambiarNumeracion && (int)this.txtNumEstado.Value == 6) this.txtNumero.Enabled = true;			
						
						this.txtNotas.Enabled = true;

						#region Origen Factura 
						if (idTipoOrigen == 0 || idTipoOrigen == 43)
						{
							this.txtSerial.Enabled = true;
							this.grdSerial.Enabled = true;
						}

						if (idTipoOrigen == 0 || idTipoOrigen == 11 || idTipoOrigen == 13 || idTipoOrigen == 44)
						{
							if ((int)this.cmbTipoRuc.Value == 4 && ((int)this.cmbFormaPago.Value != 3 || (int)this.cmbFormaPago.Value != 11))
							{
								this.btnRetener.Enabled = true;
								this.txtSerieRetencion.Enabled = true;
								this.txtAutorizacionRetencion.Enabled = true;
								this.txtNumeroRetencion.Enabled = true;				
								this.dtFechaRetencion.Enabled = true;

								DateTime dtFecha = (DateTime)this.dtFecha.Value;

								this.dtFechaRetencion.Value = (DateTime)this.dtFecha.Value;
								this.dtFechaRetencion.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
								this.dtFechaRetencion.CalendarInfo.MaxDate = dtFecha.AddDays(5);
							}
						}

						if (idTipoOrigen == 11 || idTipoOrigen == 13 || idTipoOrigen == 44)
						{
							this.chkEntregado.Enabled = true;
							this.dtFechaEntrega.Enabled = true;
							this.dtpHoraEntrega.Enabled = true;
							this.cmbDirecciones.Enabled = true;
							this.txtNotas.Enabled = true;					
						}

						if (idTipoOrigen == 43)
						{
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieL"].CellActivation = Activation.AllowEdit;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.Disabled;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.Disabled;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.Disabled;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = Activation.Disabled;
							this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = Activation.Disabled;
						}
						#endregion Origen Factura 

						ModificaGrilla(true);
						ModificaGrillaSerial(true);

						this.btnNuevo.Enabled = false;
						this.btnConsultar.Enabled = false;
						this.btnEditar.Enabled = false;
						this.btnAnular.Enabled = false;
						this.btnImprimir.Enabled = false;
						this.btnGuardar.Enabled = true;
						this.btnImprimeDctos.Enabled = false;
						if ((int)this.txtNumEstado.Value == 6) this.btnTemporal.Enabled = true;
						this.btnAsiento.Enabled = false;
						this.btnBuscar.Enabled = false;						
						if (miAcceso.CambiarDescuento && (int)this.txtNumEstado.Value == 6 && !bCrediPoint) this.btnClaveDescuento.Enabled = true;

						this.cmbVendedor.Enabled = true;
						this.txtNumeroCredito.Enabled = true;
	
						if (miAcceso.BModificarBodega) this.cmbBodega.Enabled = true; else this.cmbBodega.Enabled = false;

						if (miAcceso.BModifcarFormaDePago) 
						{
							if ((int)this.cmbFormaPago.Value != 6 || (int)this.cmbFormaPago.Value != 9) this.cmbFormaPago.Enabled = true;
						}
						else this.cmbFormaPago.Enabled = false;

						#region Anticipos
						if (((int)this.cmbFormaPago.Value == 9 && (decimal)this.txtCuotaDeEntrada.Value > 0))
						{
							dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
							if (dAnticipo > 0)
							{
								if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
								{
									bUsoAnticipo = true;
								}
								else
								{
									bUsoAnticipo = false;
									dAnticipo = 0;
								}
							}
						}
						else
						{
							dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
							if (dAnticipo > 0)
							{
								if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
								{
									bUsoAnticipo = true;
								}
								else
								{
									bUsoAnticipo = false;
									dAnticipo = 0;
								}
							}
						}
						#endregion Anticipos

						if (BRetencion)
						{
							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) == 0)
							{
								this.txtSerieRetencion.Enabled = true;
								this.txtNumeroRetencion.Enabled = true;
								this.txtAutorizacionRetencion.Enabled = true;
						
								this.btnRetener.Enabled = true;
								this.btnCancelarRetencion.Enabled = true;
							}
						}
						else
						{
							if ((int)this.cmbTipoRuc.Value == 4 && ((int)this.cmbFormaPago.Value != 3 && (int)this.cmbFormaPago.Value != 11)) EstadoRetencion(true); else EstadoRetencion(false);
						}
			
						this.txtNotas.Enabled = true;
						this.txtSerial.Enabled = true;
						this.chkEntregado.Enabled = true;

						if (this.chkEntregado.Checked)
						{
							this.dtFechaEntrega.Enabled = true;
							this.dtpHoraEntrega.Enabled = true;
							this.cmbDirecciones.Enabled = true;
						}

						if ((int)this.txtNumEstado.Value == 6)
						{
							SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
							dr.Read();
							if (dr.HasRows)
							{
								this.txtSerie.Text = dr.GetValue(0).ToString();
								this.txtAutFactura.Text = dr.GetValue(1).ToString();
								if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
								bNumeracionAutomatica = dr.GetBoolean(3);
							}
							dr.Close();

							if (!bNumeracionAutomatica)	this.txtNumero.Enabled = true; else this.txtNumero.Enabled = false;
						}

						if ((int)this.txtNumEstado.Value == 6) bNuevo = true; else bNuevo = false;
					
						if (this.miAcceso.CambiarDescuento) this.btnClaveDescuento.Enabled = true;

						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
						idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

						bEdicion = true;
					}
				}
			}
		}

		private void NuevoRegistro()
		{
			this.cmbTipoRuc.Value = 5;						
			this.txtNumContadoCredito.Value = 1;	
											
			bNuevo = true;

			DateTime dtRetFecha = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;

			//if ((int)this.txtBodegaPredef.Value == 0) 
				this.cmbBodega.Value = 3;
			//else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;
						
			this.dtFechaRetencion.Value = DateTime.Today;
			this.dtFechaRetencion.CalendarInfo.MinDate = DateTime.Today;
			this.dtFechaRetencion.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today;

			#region Habilita controles 
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			else this.dtFecha.Enabled = false;
			if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;
			else this.txtNumero.Enabled = false;
//			if (miAcceso.BModificarBodega) this.cmbBodega.Enabled = true;
//			else this.cmbBodega.Enabled = false;
			//						if (miAcceso.CambiarDescuento) this.btnClaveDescuento.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnTemporal.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled = true;
			this.btnCliente.Enabled = true;
			this.btnBuscar.Enabled = false;
			this.cmbVendedor.Enabled = true;

			ModificaGrilla(true);
			ModificaGrillaSerial(true);
						
			this.txtSerial.Enabled = true;
			this.chkEntregado.Enabled = true;
			this.txtNotas.Enabled = true;
			#endregion Habilita controles 						
			bNumeracionAutomatica = false;						
			bEdicion = true;						
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			this.cmbFormaPago.Value = 9;
			
			this.NuevoRegistro();			

			idTipoPrecio = 4;// Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value));			

			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.txtNumeroCredito.Enabled = true;
			this.cmbVendedor.Focus();
		}
		
		bool bCargaPromocionFP = false;
		bool bFormaPagoPromocional = false;
		DateTime dtFecVencCupon = DateTime.Today;
		string sNumero = "";

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				//if (!bEdicion) return;

				if (bNuevo) this.txtNumero.Text = "";

				this.txtNumIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
					
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalVendedor {0}", (int)this.cmbBodega.Value));

//				if (BUsoNotaDeCredito) idTipoPrecio = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.txtNumidFPNC.Value, (int)this.cmbBodega.Value));
//				else idTipoPrecio = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value));

				#region S.R.I.
				if (bNuevo)
				{
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
					dr.Read();
					if (dr.HasRows)
					{
						this.txtSerie.Text = dr.GetValue(0).ToString();
						this.txtAutFactura.Text = dr.GetValue(1).ToString();
						if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
						//bNumeracionAutomatica = dr.GetBoolean(3);
					}
					dr.Close();

					if (!bNumeracionAutomatica)	this.txtNumero.Enabled = true;
					else this.txtNumero.Enabled = false;
				}
				#endregion S.R.I.
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{			
//			if (bNuevo && (int)this.txtNumidOrigen.Value == 0)
//			{
			if (bEdicion)
			{
				e.Row.Cells["idDetCompra"].Value = 0;
				e.Row.Cells["idOrigen"].Value = 0;
				
				e.Row.Cells["Bodega"].Value = (int) this.cmbBodega.Value;
				e.Row.Cells["Signo"].Value = IdSigno;
				e.Row.Cells["Manual"].Value = false;
				e.Row.Cells["Tipo"].Value = 0;
				e.Row.Cells["Principal"].Value = false;
				e.Row.Cells["Entregado"].Value = false;
				//if (this.ultraGrid1.Rows.Count == 1) 
				e.Row.Cells["Posicion"].Value = this.ultraGrid1.Rows.Count;
				e.Row.Cells["Bloqueado"].Value = this.ultraGrid1.Rows.Count;
				e.Row.Cells["idCombo"].Value = 0;
				//else e.Row.Cells["Posicion"].Value = this.ultraGrid1.Rows.Count + 1;
			}
//			}
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
		/*	int iTotalLineas = ultraGrid1.Rows.Count;
			string stExec = string.Format("Exec CompraNumeroLineas 1, {0}", iTotalLineas);
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);

			if (stMensaje.Length > 0) 
			{
				e.Cancel = true;
				MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information); // "No puede crear más lineas");
			}*/
		}
		
		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion) 
			{
				if (!bEliminaAlValidar)
				{
					if (!bFormaPagoPromocional)
					{
						#region Validacion
//						if ((int)this.cmbFormaPago.Value == 9)
//						{
//							MessageBox.Show("No puede eliminar filas de Facturas de CREDIPOINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							e.Cancel = true;
//							return;
//						}

						if (idTipoOrigen == 11)
						{
							MessageBox.Show("No puede eliminar filas de Facturas de RESERVAS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
							return;
						}

						if (idTipoOrigen == 13)
						{
							MessageBox.Show("No puede eliminar filas de Facturas de PEDIDOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
							return;
						}
						#endregion Validacion			

						if (!bElCombo)
						{					
							if (!bNuevo && !miAcceso.BEliArt) 
							{
								MessageBox.Show("No puede eliminar filas de Facturas ya registradas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}
				
							int VerifCantArticulo = 0;

							for (int i=0; i<e.Rows.Length; i++)
							{
								if (e.Rows[i].Cells["idArticulo"].Value == System.DBNull.Value)
								{
									e.DisplayPromptMsg = false;	
								}
								else
								{
									if ((bool)e.Rows[i].Cells["Principal"].Value && !miAcceso.BEliArtCmbProm)
									{
										MessageBox.Show("No puede Eliminar el Producto Principal de un Combo o Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										e.Cancel = true;
										return;
									}

									int idArticuloElimina = 0;
									if (e.Rows[i].Cells["idArticulo"].Value != System.DBNull.Value) idArticuloElimina = (int)e.Rows[i].Cells["idArticulo"].Value;
									int idComboElimina = 0;
									if (e.Rows[i].Cells["idCombo"].Value != System.DBNull.Value) idComboElimina = (int)e.Rows[i].Cells["idCombo"].Value;

									if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Bloqueado Bit = 0 Select @Bloqueado = Bloqueado From ComboDetalle Where (idArticulo = {0} Or idArticuloAlterno1 = {0} Or idArticuloAlterno2 = {0} Or idArticuloAlterno3 = {0} Or idArticuloAlterno4 = {0} Or idArticuloAlterno5 = {0}) And idCombo = {1} Select @Bloqueado", idArticuloElimina, idComboElimina), true))
									{
										MessageBox.Show("Este Articulo no puede ser Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										e.Cancel = true;
										return;							 
									}

									if ((int)e.Rows[i].Cells["Tipo"].Value > 0 && !miAcceso.BEliArtCmbProm)
									{
										MessageBox.Show("No puede Eliminar un Producto de un Combo o Promoción \n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										e.Cancel = true;
										return;
									}
					
//									if (!bNuevo && !miAcceso.BEliArt && (int)this.cmbFormaPago.Value == 9)
//									{
//										MessageBox.Show("No puede Eliminar un Producto de CrediPoint Aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//										e.Cancel = true;
//										return;
//									}

									string stMensaje = string.Format("Select ISNULL(ManejaSerial, 0) From Articulo Where idArticulo = {0}", (int)e.Rows[i].Cells["idArticulo"].Value);
									bool bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);

									if (bManejaSerial == true)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
										{
											int IdArtSerie = (int) dr1.Cells["idArticulo"].Value;

											if (IdArtSerie == (int)e.Rows[i].Cells["idArticulo"].Value) VerifCantArticulo++;
										}
									}
								}
							}
          
							if (VerifCantArticulo > 0)
							{
								MessageBox.Show(string.Format("Este Artículo {0} tiene seriales registradas, \n\n Elimine primero las seriales", VerifCantArticulo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}
							#region Elimina
							if (DialogResult.No ==	MessageBox.Show("¿Esta Seguro de Borrar este Artículo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;					
							else
							{
								if (bNuevo)
								{							
									e.DisplayPromptMsg = false;	
									if ((bool)this.ultraGrid1.ActiveRow.Cells["Principal"].Value)
									{
										idComboElimina = (int)this.ultraGrid1.ActiveRow.Cells["idCombo"].Value;
										bElCombo = true;
									}
									if ((int)this.txtNumEstado.Value == 6)
									{
										for (int i=0; i<e.Rows.Length; i++)
										{
											string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value);
											Funcion.EjecutaSQL(cdsSeteoF, sSQLBorraLinea, true);
										}
									}
								}
								else
								{
									for (int i=0; i<e.Rows.Length; i++)
									{
										string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLBorraLinea, true);

										this.btnCancelar.Enabled = false;
									}
								}
							}
							#endregion Elimina
						}
						else e.DisplayPromptMsg = false;
					}
				}
				else e.DisplayPromptMsg = false;
			}
			else e.Cancel = true;
		}

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true;
				string stError = e.DataErrorInfo.InvalidValue.ToString();
				string stCol = e.DataErrorInfo.Cell.Column.ToString();
				string stMensaje = "Error Detalle: VENTAS";
				
				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);

				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol + "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", 
					"Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void grdSerial_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["Estado"].Value == 11)
			{
				e.Row.Appearance.BackColor = Color.Red;
				e.Row.Appearance.BackColor2 = Color.Red;
			}
		}

		private void grdSerial_DoubleClick(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.grdSerial.ActiveRow.Cells["idSerial"].Value != System.DBNull.Value) 
				{
					string Serial = this.grdSerial.ActiveRow.Cells["Serial"].Value.ToString();

					if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Desea modificar la serial '{0}'?", Serial), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						this.txtSerial.Text = Serial; 
						ModSerial = true;
					}
					else ModSerial = false;
				}
			}
		}

		private void txtSerieRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumeroRetencion.Focus();
		}

		private void txtNumeroRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.dtFechaRetencion.Focus();
		}

		private void txtAutorizacionRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtSerieRetencion.Focus();
		}

		private void btnRetener_Click(object sender, System.EventArgs e)
		{
			dValorRetencion = 0.00m;
			
			if ((decimal)this.txtTotal.Value == 0.00m)
			{
				MessageBox.Show("No puede crear la retención en 0.00", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
			}
			else if (this.txtSerieRetencion.Text.ToString() == "")
			{
				MessageBox.Show("Ingrese la serie para la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerieRetencion.Focus();				
			}
			else if (this.txtSerieRetencion.Text.ToString().Length < 6)
			{
				MessageBox.Show("La serie debe de tener 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerieRetencion.Focus();				
			}
			else if (this.txtAutorizacionRetencion.Text.ToString() == "")
			{
				MessageBox.Show("Ingrese el número de Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtAutorizacionRetencion.Focus();				
			}
			else if (this.txtAutorizacionRetencion.Text.Length < 10 || this.txtAutorizacionRetencion.Text.Length > 37)
			{
				MessageBox.Show("Longitud de la Autorización debe estar entre 10 y 37 caracteres", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtAutorizacionRetencion.Focus();				
			}
			else if (this.txtNumeroRetencion.Text.ToString() == "")
			{
				MessageBox.Show("Ingrese el número de la retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNumeroRetencion.Focus();				
			}
			else if (this.dtFechaRetencion.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha de Retención.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNumeroRetencion.Focus();
			}		
			else
			{
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaRetDetalle -1");

				#region IVA
				int idAgente = 0;
				int idSujeto = 0;
				bool Bienes = false;
				bool Bienes100 = false;
				bool Servicios = false;
				bool Servicios100 = false;
				int idRetencionIVA = 0;
				decimal sPorc = 0;
				string sCodigo = "";
				string sDetalle = "";
				decimal dBaseIVA = 0.00m;	
				decimal dIVARetenido = 0.00m;
				int iContI = 0;
				int idArticulo = 0;
				decimal dImpuesto = 0.00m;
				decimal dSubtotal = 0.00m;
				decimal dDescuento = 0.00m;
				decimal dIVA = 0.00m;
				DateTime dtFecha = (DateTime)this.dtFechaRetencion.Value;

				idAgente = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select ISNULL(idAgenteRetencion, 0) From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value));
				idSujeto = Funcion.iEscalarSQL(cdsSeteo, "Select ISNULL(idAgenteRetencion, 0) From Seteo");
						
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					Bienes = false;
					Bienes100 = false;
					Servicios = false;
					Servicios100 = false;
					idRetencionIVA = 0;
					sPorc = 0;
					sCodigo = "";
					sDetalle = "";
					dBaseIVA = 0.00m;	
					dIVARetenido = 0.00m;
					iContI = 0;					
					idArticulo = (int)dr.Cells["idArticulo"].Value;
					dImpuesto = (decimal)dr.Cells["Impuesto"].Value;
					dSubtotal = (decimal)dr.Cells["Subtotal"].Value;
					dDescuento = (decimal)dr.Cells["DescuentoArt"].Value;
					dIVA = Math.Round(((dSubtotal - dDescuento) * dImpuesto)/100, 2);

					#region Codigo Nuevo 
					string sSQLIVA = string.Format("Exec RetencionArticuloIVA {0}, {1}, '{2}', 1, {3}", (int)dr.Cells["idArticulo"].Value, dIVA, dtFecha.ToString("yyyyMMdd"), (int)this.txtNumIdCliente.Value);
					SqlDataReader drIVA = Funcion.rEscalarSQL(cdsSeteoF, sSQLIVA, true);					
					drIVA.Read();
					if (drIVA.HasRows)
					{
						idRetencionIVA = drIVA.GetInt32(0); 
						sPorc = drIVA.GetDecimal(1);
						sCodigo = drIVA.GetString(2);
						sDetalle = drIVA.GetString(3);
						dBaseIVA = drIVA.GetDecimal(4);	
						dIVARetenido = drIVA.GetDecimal(5);	
					}
					drIVA.Close();

					#region Valida Retencion I.V.A. Existente
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid2.Rows.All)
					{
						if ((int)dr1.Cells["idRetencionIVARenta"].Value == idRetencionIVA && !(bool) dr1.Cells["IVARenta"].Value)	
						{
							dr1.Cells["Base"].Value = Convert.ToDecimal(dr1.Cells["Base"].Value) + dBaseIVA;
							dr1.Cells["Valor"].Value = Convert.ToDecimal(dr1.Cells["Valor"].Value) + dIVARetenido;
							iContI ++;
						}
					}
					#endregion Valida Retencion I.V.A. Existente

					if (idRetencionIVA > 0 && iContI == 0)
					{
						this.ultraGrid2.Rows.Band.AddNew();
						this.ultraGrid2.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
						this.ultraGrid2.ActiveRow.Cells["idRetencion"].Value = 0;
						this.ultraGrid2.ActiveRow.Cells["IVARENTA"].Value = false;
						this.ultraGrid2.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionIVA;
						this.ultraGrid2.ActiveRow.Cells["Porcentaje"].Value = sPorc;
						this.ultraGrid2.ActiveRow.Cells["Codigo"].Value = sCodigo;
						this.ultraGrid2.ActiveRow.Cells["Detalle"].Value = sDetalle;
						this.ultraGrid2.ActiveRow.Cells["Base"].Value = dBaseIVA;
						this.ultraGrid2.ActiveRow.Cells["Valor"].Value = dIVARetenido;	
					}
					#endregion Codigo Nuevo					
				}
				#endregion IVA

				#region Renta
				#region Variables
				int idRetencionRenta = 0;
				decimal sPorcR = 0.00m;
				string sCodigoR = "";
				string sDetalleR = "";
				decimal dBaseRenta = 0.00m;
				decimal dRentaRetenida = 0.00m;				
				decimal dRenta = 0.00m;
				int iContR = 0;
				#endregion Variables

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					#region Asignacion de Variables
					idRetencionRenta = 0;
					sPorcR = 0;
					sCodigoR = "";
					sDetalleR = "";
					dBaseRenta = 0.00m;
					dRentaRetenida = 0.00m;
					dSubtotal = 0.00m;
					dDescuento = 0.00m;
					dRenta = 0.00m;
					dSubtotal = (decimal)dr.Cells["Subtotal"].Value;
					dDescuento = (decimal)dr.Cells["DescuentoArt"].Value;
					dRenta = Math.Round(dSubtotal, 2) - Math.Round(dDescuento, 2);
					iContR = 0;
					#endregion Asignacion de Variables

					#region Valores Retencion 
					string sSQLRenta = string.Format("Exec RetencionArticuloRenta {0}, {1}", (int)dr.Cells["idArticulo"].Value, dRenta);
					SqlDataReader drRenta = Funcion.rEscalarSQL(cdsSeteoF, sSQLRenta, true);
					drRenta.Read();
					if (drRenta.HasRows)
					{
						idRetencionRenta = drRenta.GetInt32(0);
						sPorcR = drRenta.GetDecimal(1);
						sCodigoR = drRenta.GetString(2);
						sDetalleR = drRenta.GetString(3);
						dBaseRenta = drRenta.GetDecimal(4);
						dRentaRetenida = drRenta.GetDecimal(5);	
					}
					drRenta.Close();
					#endregion Valores Retencion 

					#region Valida Retencion Existente
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid2.Rows.All)
					{
						if ((int)dr1.Cells["idRetencionIVARenta"].Value == idRetencionRenta && (bool) dr1.Cells["IVARenta"].Value)	
						{
							dr1.Cells["Base"].Value = Convert.ToDecimal(dr1.Cells["Base"].Value) + dBaseRenta;
							dr1.Cells["Valor"].Value = Convert.ToDecimal(dr1.Cells["Valor"].Value) + dRentaRetenida;
							iContR ++;
						}
					}
					#endregion Valida Retencion Existente

					#region Asignacion Retencion
					if (iContR == 0 && idRetencionRenta > 0)
					{
						this.ultraGrid2.Rows.Band.AddNew();
						this.ultraGrid2.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
						this.ultraGrid2.ActiveRow.Cells["idRetencion"].Value = 0;
						this.ultraGrid2.ActiveRow.Cells["IVARENTA"].Value = true;
						this.ultraGrid2.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetencionRenta;
						this.ultraGrid2.ActiveRow.Cells["Porcentaje"].Value = sPorcR;
						this.ultraGrid2.ActiveRow.Cells["Codigo"].Value = sCodigoR;
						this.ultraGrid2.ActiveRow.Cells["Detalle"].Value = sDetalleR;
						this.ultraGrid2.ActiveRow.Cells["Base"].Value = dBaseRenta;
						this.ultraGrid2.ActiveRow.Cells["Valor"].Value = dRentaRetenida;
					}
					#endregion Asignacion Retencion
				}
				#endregion Renta

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				{
					dValorRetencion = dValorRetencion + Convert.ToDecimal(dr.Cells["Valor"].Value);
				}

				BRetencion = true;
			}
		}

		private void txtNumeroRetencion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumeroRetencion.Text.ToString().Trim().Length > 0)
			{
				char Pad = '0';
				this.txtNumeroRetencion.Text = this.txtNumeroRetencion.Text.ToString().PadLeft(9, Pad);
			}
		}

		private void txtSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnCancelarRetencion_Click(object sender, System.EventArgs e)
		{
			this.txtidRetencion.Value = 0;
			this.txtIdAsientoRetencion.Value = 0;
			this.txtSerieRetencion.Text = "";
			this.txtAutorizacionRetencion.Text = "";
			this.txtNumeroRetencion.Text = "";
			this.dtFechaRetencion.Value = System.DBNull.Value;
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaRetDetalle -1");
			BRetencion = false;														
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Valida identificación
			if (bEdicion)
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.cmbTipoRuc.Value);
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;			
				}
			}
			#endregion Valida identificación
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtNumIdCliente.Value = 0;
				this.txtNombre.Text = "";

				#region Busca Cliente 
				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, false);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drC.Read();
				if (drC.HasRows)
				{
					this.txtNumIdCliente.Value = (int) drC.GetValue(0);
					this.txtRuc.Text = (string) drC.GetValue(1);
					this.txtNombre.Text = (string) drC.GetValue(2);

					drC.Close();
				}
				else
				{						
					using (VentaCliente miVCliente = new VentaCliente(this.txtRuc.Text, 0, (int)this.cmbTipoRuc.Value))
					{
						if (DialogResult.OK == miVCliente.ShowDialog())
						{
							MessageBox.Show("Cliente registrado correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.cmbTipoRuc.Value = (int)miVCliente.cmbTipoRuc.Value;
							this.txtNumIdCliente.Value = (int) miVCliente.txtidCliente.Value;
							this.txtRuc.Text = miVCliente.txtIdentificacion.Value.ToString().Trim();
							this.txtNombre.Text = miVCliente.txtNombre.Value.ToString().Trim();

							if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
							if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;
						}
						else this.txtRuc.Focus();
					}			
				}
				drC.Close();		

				#region Anticipos
				if (((int)this.cmbFormaPago.Value == 9 && (decimal)this.txtCuotaDeEntrada.Value > 0))
				{
					dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
					if (dAnticipo > 0)
					{
						if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						{
							bUsoAnticipo = true;
						}
						else
						{
							bUsoAnticipo = false;
							dAnticipo = 0;
						}
					}
				}
				else
				{
					dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
					if (dAnticipo > 0)
					{
						if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						{
							bUsoAnticipo = true;
						}
						else
						{
							bUsoAnticipo = false;
							dAnticipo = 0;
						}
					}
				}
				#endregion Anticipos

				#region Valida crea retención
				if ((int)this.cmbTipoRuc.Value == 4 && ((int)this.cmbFormaPago.Value != 3 && (int)this.cmbFormaPago.Value != 11)) EstadoRetencion(true); else EstadoRetencion(false);				
				#endregion Valida crea retención
				#endregion Busca Cliente
			}
		}

		private void chkEntregado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFechaEntrega.Focus();
		}

		private void dtFechaEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtpHoraEntrega.Focus();
		}

		private void dtpHoraEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbDirecciones.Focus();
		}

		private void cmbDirecciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnGuardar.Focus();
		}

		private void Consultar(int IdCompra)
		{
			try
			{	
				#region Botones
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnTemporal.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Text = "Anular";
				this.btnCliente.Enabled = false;
				#endregion Botones

				#region Maestro
				#region Orden de Campos 
				/*c.idCompra, c.idOrigen, c.idSucursal, c.idTipoFactura, c.idComprobante, c.Estado, 0-5
				Case c.Borrar When 0 Then Case c.Estado When 6 Then 'TEMPORAL' When 10 Then 'CONTABILIZADO' When 11 Then 'MANUAL' End When 1 Then 'ANULADO' End, 6
				c.Consignacion, c.ContadoCredito, c.idFormaPago, c.Bodega, c.idProyecto, c.idVendedor, c.Numero, c.Comprobante, 7-14
				c.idTipoRuc, c.idCliente, cl.Ruc, cl.Nombre, c.Fecha, c.SerieFactura, c.AutFactura, 15-21
				ISNULL(c.idEntidadFinanciera, 0), c.FechaCaducidad, c.Entregado, c.FechaEntrega, c.HoraEntrega, c.idDireccion, c.Notas, 22-28
				c.CrediPoint, c.TotalCuotas, c.NumCuotas, c.CuotaEntrada, FechaPrimerPago, 29-33
				c.SubtotalExcento, c.Descuento0, c.SubtotalIva, c.Descuento, c.Iva, c.Total, 34-39
				c.Impreso, c.idAsiento, c.Borrar, c.Usuario, c.idProyectoZoom, pz.Nombre, 40-45
				c.CuotaAsignada, c.Interes, c.CuotaDisponible, c.CuotaUsada, c.Procesado, c.ValorCuotas 46-51 */
				#endregion Orden de Campos 

				this.txtNumIdCompra.Value = IdCompra;
				int iBodega = 0;
				
				string sSQL = string.Format("Exec VentaConsultaIndividual {0}", (int)this.txtNumIdCompra.Value);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{					
					this.txtNumidOrigen.Value = dr.GetInt32(1);
					this.txtIdSucursal.Value  = dr.GetInt32(2);
					this.txtIdTipoFactura.Value = 1;
					this.txtIdComprobante.Value  = dr.GetInt32(4);
					this.txtNumEstado.Value = dr.GetInt32(5);
					this.lblEstado.Text = dr.GetString(6);
					bConsignacion = dr.GetBoolean(7);
					this.txtNumContadoCredito.Value = dr.GetInt32(8);
					this.cmbFormaPago.Value = dr.GetInt32(9);
					iBodega = dr.GetInt32(10);
					this.txtNumIdProyecto.Value = dr.GetInt32(11);
					this.cmbVendedor.Value = dr.GetInt32(12);
					this.txtNumero.Text = dr.GetValue(13).ToString();
					this.txtComprobante.Text = dr.GetValue(14).ToString();
					this.cmbTipoRuc.Value = dr.GetInt32(15);
					this.txtNumIdCliente.Value = dr.GetInt32(16);
					this.txtRuc.Text = dr.GetValue(17).ToString();
					this.txtNombre.Text = dr.GetValue(18).ToString();
					if ((int)this.txtNumEstado.Value != 6) this.dtFecha.Value = dr.GetDateTime(19); else this.dtFecha.Value = DateTime.Today;
					this.txtSerie.Text = dr.GetValue(20).ToString();
					this.txtAutFactura.Text = dr.GetValue(21).ToString();
					this.txtIdEF.Value = dr.GetInt32(22);
					if (dr.GetValue(23) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(23);
					this.chkEntregado.Checked = dr.GetBoolean(24);

					if (this.chkEntregado.Checked)
					{						
						if (dr.GetValue(25) != System.DBNull.Value) this.dtFechaEntrega.Value = dr.GetDateTime(25);
						if (dr.GetValue(26) != System.DBNull.Value) this.dtpHoraEntrega.Value = dr.GetDateTime(26);
						this.cmbDirecciones.Value = dr.GetInt32(27);
					}

					this.txtNotas.Text = dr.GetValue(28).ToString();				
					
					bCrediPoint = dr.GetBoolean(29);

					if (bCrediPoint)
					{						
						dTotalCuotas = dr.GetDecimal(30);
						NumCuotas = dr.GetInt32(31);
						this.txtCuotaDeEntrada.Value = dr.GetDecimal(32);
						if (dr.GetValue(33) != System.DBNull.Value) dtFechPrimPago = dr.GetDateTime(33);
					}

					this.txtIva0.Value = dr.GetDecimal(34);
					this.txtDescIva0Total.Value = dr.GetDecimal(35);
					this.txtIva.Value = dr.GetDecimal(36);
					this.txtDescIvaTotal.Value = dr.GetDecimal(37);
					this.txtIvaTotal.Value = dr.GetDecimal(38);
					this.txtTotal.Value = dr.GetDecimal(39);

					bImpreso = dr.GetBoolean(40);
					this.txtNumIdAsiento.Value = dr.GetInt32(41);
					bAnulado = dr.GetBoolean(42);
					this.lblUsuario.Text = dr.GetString(43);

					this.txtNumIdproyectoZoom.Value = dr.GetInt32(44);
										
					dCuotaAsignada = dr.GetDecimal(46);
					dInteres = dr.GetDecimal(47);
					dCuotaDisponible = dr.GetDecimal(48);
					dCuotaUsada = dr.GetDecimal(49);
					AutorizaEntrada = dr.GetBoolean(50);
					dValorCuota = dr.GetDecimal(51);
					this.txtidBanco.Value = dr.GetInt32(52);
					this.txtidTarjeta.Value = dr.GetInt32(53);
					this.txtidPlan.Value = dr.GetInt32(54);
					this.txtidPlazo.Value = dr.GetInt32(55);
				}
				dr.Close();

				this.txtBaseImponible.Value = (decimal)this.txtIva.Value - (decimal)this.txtDescIvaTotal.Value;

				this.cmbBodega.Value = iBodega;

				this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", (int)this.txtNumIdCliente.Value));
				#endregion Maestro

//				if (!bAnulado) 
//				{
//					string sSQLAP = string.Format("Exec ActualizaPosicion {0}", (int)this.txtNumIdCompra.Value);
//					Funcion.EjecutaSQL(cdsSeteoF, sSQLAP, true);
//				}

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value));
				this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value));

				if ((int)this.cmbFormaPago.Value == 3) this.lblPlan.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Descripcion From Planes Where idPlan = {0}", (int)this.txtidPlan.Value));				

				#region Retenciones
				int idTipoRuc = 0;
				if (this.cmbTipoRuc.ActiveRow != null) idTipoRuc = (int)this.cmbTipoRuc.Value;

				if (idTipoRuc == 4)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Retenciones Where idCompra = {0} And ClienteProveedor = 0 And Estado <> 2", this.txtNumIdCompra.Value)) > 0)
					{
						string sSQLRet = string.Format("Exec ConsultaRetencionVenta {0}", (int)this.txtNumIdCompra.Value);
						SqlDataReader drRet = Funcion.rEscalarSQL(cdsSeteoF, sSQLRet, true);
						drRet.Read();						
						if (drRet.HasRows)
						{
							this.txtidRetencion.Value = drRet.GetInt32(0);
							this.txtSerieRetencion.Text = drRet.GetString(1);
							this.txtAutorizacionRetencion.Text = drRet.GetString(2);
							this.txtNumeroRetencion.Text = drRet.GetString(3);
							this.dtFechaRetencion.Value = drRet.GetDateTime(4);
							this.txtIdAsientoRetencion.Value = drRet.GetInt32(5);
						}
						drRet.Close();
								
						this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaRetDetalle {0}", (int)this.txtNumIdCompra.Value));		
						
						BRetencion = true;
					}
				}
				#endregion Retenciones

				if (!bAnulado)
				{
					if ((int)this.txtNumEstado.Value != 6)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
						{
							if (miAcceso.Editar) this.btnEditar.Enabled = true;// && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) 
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							{
								if ((int)this.cmbFormaPago.Value != 9)
								{
									if ((DateTime)this.dtFecha.Value == DateTime.Today)this.btnAnular.Enabled = true;
								}
							}
							if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						}
					}
					else
					{
						if (miAcceso.Editar && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) this.btnEditar.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (miAcceso.BAnulaSoloDelDia && !bAnulado)
						{
							if ((int)this.cmbFormaPago.Value != 9)
							{
								if ((DateTime)this.dtFecha.Value == DateTime.Today)this.btnAnular.Enabled = true;
							}
						}
						if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
						//if (miAcceso.CrearAsiento && !bAnulado) this.btnAsiento.Enabled = true;
					}

					if ((int)this.txtNumEstado.Value != 6)
					{
						if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
						if (miAcceso.CrearAsiento && !bAnulado) this.btnAsiento.Enabled = true;
					}

					if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
					
					if (bCrediPoint) this.btnImprimeDctos.Enabled = true;					
				}
				else if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;

				this.btnCancelar.Enabled = true;

				if (bAnulado) this.btnAnular.Text = "Eliminar";
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaFacturas miBusqueda = new frmBuscaFacturas((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					this.Consultar((int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
				}
			}		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				string sMensaje = "ANULAR";
				if (bAnulado) sMensaje = "ELIMINAR";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede '{1}' esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper(), sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}
			else 
			{
				if (bNuevo)
				{	
					#region Nuevo Registro	
					if (this.txtNumero.Text.ToString().Length == 0 && !bNumeracionAutomatica)
					{
						MessageBox.Show("Ingrese el número de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumero.Focus();						
					}
					else
					{
						string sSQLVNE = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)this.txtNumIdCompra.Value, this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
						string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVNE, true);

						if (sMensaje.ToString().Length > 0 && !bNumeracionAutomatica)
						{
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtNumero.Focus();					
						}
						else
						{
							using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
							{
								if (DialogResult.OK == miNota.ShowDialog())
								{
									#region Maestro
									this.lblEstado.Text = "ANULADO";
									bAnulado = true;
									int idDireccion = 0;
									int idVendedor = 0;
									if (this.cmbDirecciones.Text.ToString().Trim().Length > 0) idDireccion = (int)this.cmbDirecciones.Value;
									if (this.cmbVendedor.Text.ToString().Trim().Length > 0) idVendedor = (int)this.cmbVendedor.Value;

									#region Fechas
									DateTime dtFecha = (DateTime)this.dtFecha.Value;
									DateTime dtFechaCad = DateTime.Today;
									DateTime dtFechaEnt = DateTime.Today;
									DateTime dtHoraEntrega = DateTime.Today;
									bool bFecCad = false;
									bool bFecEnt = false;
									bool bHorEnt = false;
									bool bFecPrimPag = false;
									if (this.cmbCaducidad.Value != System.DBNull.Value) 
									{
										bFecCad = true;
										dtFechaCad = (DateTime) this.cmbCaducidad.Value;
									}
									if (this.dtFechaEntrega.Value != System.DBNull.Value) 
									{
										bFecEnt = true;
										dtFechaCad = (DateTime) this.dtFechaEntrega.Value;
									}
									if ((bool)this.chkEntregado.Checked)
									{
										if (this.dtpHoraEntrega.Value != System.DBNull.Value) 
										{
											bHorEnt = true;
											dtFechaCad = (DateTime) this.dtpHoraEntrega.Value;
										}
									}
									if (bCrediPoint) bFecPrimPag = true;
									#endregion Fechas

									if (bNumeracionAutomatica) this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 1, {0}, 0", (int)this.cmbBodega.Value)); //Funcion.sEscalarSQL(cdsSeteoF, "EXEC NumeroFacturaEstacion 1"); 

									this.txtNotas.Text = miNota.txtNotas.Text;
									bAnulado = true;

									string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, 0.0, 0.0, 0.0, 0.0, 0, 0.0, 0, 0, 0, 0, 0, ''", 
										(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value,
										(int)this.txtIdSucursal.Value, 
										(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
										(int)this.txtNumEstado.Value,
										bConsignacion,
										(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
										(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
										idVendedor,
										this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
										(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
										dtFecha.ToString("yyyyMMdd"), 
										this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), 
										(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion,
										this.txtNotas.Text.ToString(),
										bCrediPoint, dTotalCuotas, NumCuotas, (decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"),
										(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
										bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado);
									this.txtNumIdCompra.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
									#endregion Maestro

									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago 0, {0}, 1, 1, 0, 0, 0, '', '', '', '', '', 0, '', '', 0, 0, 0, 0", (int)this.txtNumIdCompra.Value);							
									Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);

									#region Controles
									this.cmbFormaPago.Enabled = false;
									this.cmbBodega.Enabled = false;
									this.cmbVendedor.Enabled = false;
									this.cmbTipoRuc.Enabled = false;

									this.txtNumero.Enabled = false;
									this.txtRuc.Enabled = false;
									this.dtFecha.Enabled = false;

									if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
									if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
									if (miAcceso.Editar) this.btnEditar.Enabled = true;
									if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
									else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
									if (!miAcceso.Eliminar) this.btnAnular.Enabled = false;
									this.btnGuardar.Enabled = false;
									this.btnAnular.Enabled = false;
									this.btnCancelar.Enabled = true;
									this.btnClaveDescuento.Enabled = false;
									this.btnCliente.Enabled = false;
									bDescuentoActivo = false;

									bNuevo = false;
									bEdicion = false;

									ModificaGrilla(false);
									HabilitaTab(false);
									#endregion Controles

									#region Impresión
									if (bNuevo)
									{
										frmMensajeNumeración miMsje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, 0, false);
									}
									#endregion Impresión									
								}
							}
						}
					}
					#endregion Nuevo Registro
				}
				else
				{
					this.Consultar((int)this.txtNumIdCompra.Value);

					if (bAnulado)
					{
						#region Elimina Registro
						if (DialogResult.Yes == MessageBox.Show("Esta seguro de ELIMINAR este Documento \n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
						{
							string sSQLBorrar = string.Format("Exec BorrarCompraVenta {0}, {1}, {2}", (int)this.txtNumIdCompra.Value, (int)this.txtNumIdAsiento.Value, (int)this.txtIdTipoFactura.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLBorrar, true);

							MessageBox.Show("Documento Eliminado del sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.btnCancelar_Click(sender, e);
						}
						#endregion Elimina Registro
					}
					else
					{
						if (bAnulado)
						{
							MessageBox.Show("El Documento esta Anulado", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						}
						else if (!miAcceso.Anular)
						{
							MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						}
						else
						{
							bool bBloquear = false;
							bBloquear = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Bloquear, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true);
							if (bBloquear)
							{
								MessageBox.Show("No puede Anular la Factura esta Bloqueada", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
							}
							else
							{
								string sNumeroNC = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VentasValidaNotaDeCredito {0}", (int)this.txtNumIdCompra.Value));
								if (sNumeroNC.Length > 0)
								{
									MessageBox.Show(string.Format("Esta Factura esta Asociada con la Nota de Credito N. '{0}'", sNumeroNC), "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
								else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
								{
									MessageBox.Show("No puede Anular la Factura esta Asociada a una Retencion que a Afecta a Multiples Facturas, \n\nAnule la Retencion o Consulte al Administrador", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
								}
								else if ((int)this.cmbFormaPago.Value == 6 && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec VentaValidaPagos {0}", (int)this.txtNumIdCompra.Value)) > 0)
								{
									MessageBox.Show("Esta Factura ya tiene Pagos Registrados", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
								else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))
								{
									DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
									DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
									if ((miAcceso.BAnulaSoloDelDia && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Estado, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) != 6) && dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
									{
										MessageBox.Show("Solo puede Anular Facturas de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);					
									}
									else if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular este documento.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
									{
										using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
										{
											if (DialogResult.OK == Nota.ShowDialog())
											{
												try
												{
													string sSQLAnula = string.Format("Exec AnulacionCompraVenta {0}, 1, '{1}'", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString());
													Funcion.sEscalarSQL(cdsSeteoF, sSQLAnula);

													this.Consultar((int)this.txtNumIdCompra.Value);
												}
												catch(Exception Exc)
												{
													MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + this.txtNumIdCompra.Value.ToString();
			
			stReporte += ", 0";
						
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			string stFiltro = "{Compra.idCompra} = " + this.txtNumIdCompra.Value.ToString();
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{		
				if (iTipo == 0) // Muestra en pantalla
					miReporte.ShowDialog();
				if (iTipo >= 1) // Imprime una o mas copias
				{
					string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
					string stImpresora = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraReporteImpresora '" + stReporte + "'");
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
					miReporte.ImprimeRep(1);
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
				}
				if (iTipo == -1) // Archivo PDF
				{
					miReporte.ImprimeRep(2);
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + this.txtNumIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			bImpreso = true;
			if (!miAcceso.ReImprimir) this.btnImprimir.Enabled = false;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbVendedor.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbVendedor.Focus();
					e.Cancel = true;				
				}
			}
		}

		private void mnuBuscarPrefactura_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.mnuBuscar.Show(this.btnBuscar, new Point(1, 25));
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//if (e.KeyChar == 13) this.cmbVendedor.Focus();
		}

		bool bElCombo = false;
		int idComboElimina = 0;

		private void EliminaElementosCombo()
		{
			int Tipo = 1;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (bElCombo)
				{
					if (dr.Cells["Tipo"].Value != System.DBNull.Value) Tipo = (int)dr.Cells["Tipo"].Value;

					if (Tipo > 0)// && !(bool)dr.Cells["Principal"].Value) 
					{
						if (idComboElimina == (int)dr.Cells["idCombo"].Value) dr.Delete();
					}
				}			
			}
			
			bElCombo = false;
			idComboElimina = 0;

			this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
			this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];			
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			if (bElCombo) EliminaElementosCombo();

			Total();
			
			if (this.ultraGrid1.Rows.Count > 0)
			{
				int idArticulo = 0;
				if (this.ultraGrid1.Rows[0].Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)this.ultraGrid1.Rows[0].Cells["idArticulo"].Value;
				int idPromocionValor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, 0, 0, 0, 0)", (int)this.cmbBodega.Value, idArticulo, idTipoPrecio, Convert.ToDecimal(this.txtTotal.Value)));
				if (idPromocionValor > 0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
						{							
							//							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idGrupoArticulo From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value)) == 12) 
							//							{
							if ((int)dr.Cells["Tipo"].Value == 3)
							{
								if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaEliminaArticuloPromocionPorValor {0}, {1}, {2}", idPromocionValor, (int)dr.Cells["idArticulo"].Value, this.SumaPrecios())) > 0)
								{
									bEliminaAlValidar = true;
									dr.Delete();
									this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
									this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];			
									bArtPromocion = false;
									//continue;						
								}
							}
							//							}
						}
					}
				}
			}

			Total();

			bEliminaAlValidar = false;

			#region Actualiza Posicion en las Lineas
			int iPosicion = 1;
			bool bCambio = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (this.grdSerial.Rows.Count > 0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
					{
						if (((int)dr.Cells["idArticulo"].Value == (int)dr1.Cells["idArticulo"].Value) && ((int)dr.Cells["Posicion"].Value == (int)dr1.Cells["Posicion"].Value))
						{
							dr.Cells["Posicion"].Value = iPosicion;
							dr1.Cells["Posicion"].Value = iPosicion;
							iPosicion++;
							bCambio = true;
						}						
					}	
				}
				else
				{
					dr.Cells["Posicion"].Value = iPosicion;
					iPosicion++;
					bCambio = true;
				}

				if (!bCambio)
				{
					dr.Cells["Posicion"].Value = iPosicion;
					iPosicion++;				
				}

				bCambio = false;
			}
			#endregion Actualiza Posicion en las Lineas
		}
		
		private void btnCliente_Click(object sender, System.EventArgs e)
		{	
			if ((int)this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
			}
			else
			{
				VentaCliente ClienteRapido;			
				ClienteRapido = new VentaCliente(this.txtNombre.Text.ToString(), (int) this.txtNumIdCliente.Value, (int)this.cmbTipoRuc.Value);
				ClienteRapido.ShowDialog();
			}
		}

		private void txtAutorizacionRetencion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void PedidoReserva(int idTipoFacturaD)
		{
			using (frmBuscaFacturas miBusqueda = new frmBuscaFacturas(idTipoFacturaD, 1, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{	
						#region Maestro
						idTipoOrigen = idTipoFacturaD;
						this.txtNumIdCompra.Value = 0;
						this.txtNumidOrigen.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
						this.txtIdSucursal.Value  = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idSucursal"].Value;
						this.txtIdTipoFactura.Value = 1;
						this.txtIdComprobante.Value  = 15;//(int)miBusqueda.grdFacturas.ActiveRow.Cells["idComprobante"].Value;
						this.txtNumEstado.Value = 10;
						this.lblEstado.Text = "";
						bConsignacion = (bool)miBusqueda.grdFacturas.ActiveRow.Cells["Consignacion"].Value;
						this.txtNumContadoCredito.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["ContadoCredito"].Value;
						this.cmbFormaPago.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idFormaPago"].Value;
						this.cmbBodega.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["Bodega"].Value;
						this.txtNumIdProyecto.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idProyecto"].Value;
						this.cmbVendedor.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idVendedor"].Value;
						this.txtNumero.Text = "";
						this.txtComprobante.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Numero"].Value.ToString();
						this.cmbTipoRuc.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idTipoRuc"].Value;
						this.txtNumIdCliente.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCliente"].Value;
						this.txtRuc.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Ruc"].Value.ToString();
						this.txtNombre.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Nombre"].Value.ToString();
						this.dtFecha.Value = DateTime.Today;// DateTime.Today;
						this.txtSerie.Text = miBusqueda.grdFacturas.ActiveRow.Cells["SerieFactura"].Value.ToString();
						this.txtAutFactura.Text = miBusqueda.grdFacturas.ActiveRow.Cells["AutFactura"].Value.ToString();
						if (miBusqueda.grdFacturas.ActiveRow.Cells["FechaCaducidad"].Value != System.DBNull.Value) this.cmbCaducidad.Value = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["FechaCaducidad"].Value;
						if ((bool)miBusqueda.grdFacturas.ActiveRow.Cells["Entregado"].Value)
						{
							this.chkEntregado.Checked = (bool)miBusqueda.grdFacturas.ActiveRow.Cells["Entregado"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["FechaEntrega"].Value != System.DBNull.Value) this.dtFechaEntrega.Value = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["FechaEntrega"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["HoraEntrega"].Value != System.DBNull.Value) this.dtpHoraEntrega.DateTime = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["HoraEntrega"].Value;
							this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", (int)this.txtNumIdCliente.Value));
							this.cmbDirecciones.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idDireccion"].Value;
						}
						this.txtNotas.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Notas"].Value.ToString();
						this.txtIva0.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["SubtotalExcento"].Value;
						this.txtDescIva0Total.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Descuento0"].Value;
						this.txtIva.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["SubtotalIva"].Value;
						this.txtDescIvaTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Descuento"].Value;
						this.txtIvaTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Iva"].Value;
						this.txtTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Total"].Value;
						bImpreso = false;
						this.txtNumIdAsiento.Value = 0;
						bAnulado = false;
						#endregion Maestro

						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePedidosReservasConsignaciones {0}", (int)this.txtNumidOrigen.Value));	

						this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec BuscaSerialesPedido {0}", (int)this.txtNumidOrigen.Value));
						
						#region Anticipos
						dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
						if (dAnticipo > 0)
						{
							if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
							{
								bUsoAnticipo = true;
							}
							else
							{
								bUsoAnticipo = false;
								dAnticipo = 0;
							}
						}
						#endregion Anticipos

						this.chkEntregado.Enabled = true;
						this.dtFechaEntrega.Enabled = true;
						this.dtpHoraEntrega.Enabled = true;
						this.cmbDirecciones.Enabled = true;
						this.txtNotas.Enabled = true;

						if ((int)this.cmbTipoRuc.Value == 4)
						{
							this.btnRetener.Enabled = true;
							this.txtSerieRetencion.Enabled = true;
							this.txtAutorizacionRetencion.Enabled = true;
							this.txtNumeroRetencion.Enabled = true;				
							this.dtFechaRetencion.Enabled = true;
						}
						if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
						if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;	

						//this.cmbVendedor.Enabled = true;

						this.btnNuevo.Enabled = false;
						this.btnConsultar.Enabled = false;
						this.btnEditar.Enabled = false;
						this.btnAnular.Enabled = false;
						this.btnImprimir.Enabled = false;
						this.btnGuardar.Enabled = true;
						this.btnCancelar.Enabled = true;
            this.btnBuscar.Enabled = false;

						bNuevo = true;
						bEdicion = true;
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}		
		}

		private void mnuBuscarPedido_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(13);
		}

		private void mnuBuscarReserva_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(11);
		}

		private void btnTemporal_Click(object sender, System.EventArgs e)
		{
			#region Factura Temporal
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{			
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Grabar esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}	
			else if (this.cmbVendedor.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbVendedor.Focus();				
			}			
			else if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
			}			
			else 
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.cmbTipoRuc.Value);
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();			
				}
				else
				{
					string sSQLValCl = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, false);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLValCl, true);
					drC.Read();
					if (!drC.HasRows)
					{
						drC.Close();
						MessageBox.Show("Cliente no Existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();						
					}
					else
					{
						drC.Close();			
						if (this.ultraGrid1.Rows.Count == 0)
						{
							MessageBox.Show("Ingrese los Productos para Crear la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
						}			
						else
						{				
							this.lblEstado.Text = "TEMPORAL";
							this.txtNumEstado.Value = 6;			

							#region Graba Factura
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
			
									#region Graba Factura
									#region Maestro
									int idDireccion = 0;
									if (this.cmbDirecciones.Text.ToString().Trim().Length > 0) idDireccion = (int)this.cmbDirecciones.Value;

									#region Fechas
									DateTime dtFecha = (DateTime)this.dtFecha.Value;
									DateTime dtFechaCad = DateTime.Today;
									DateTime dtFechaEnt = DateTime.Today;
									DateTime dtHoraEntrega = DateTime.Today;
									bool bFecCad = false;
									bool bFecEnt = false;
									bool bHorEnt = false;
									bool bFecPrimPag = false;
									if (this.cmbCaducidad.Value != System.DBNull.Value) 
									{
										bFecCad = true;
										dtFechaCad = (DateTime) this.cmbCaducidad.Value;
									}
									if (this.dtFechaEntrega.Value != System.DBNull.Value) 
									{
										bFecEnt = true;
										dtFechaCad = (DateTime) this.dtFechaEntrega.Value;
									}
									if ((string)this.dtpHoraEntrega.Value == "") 
									{
										bHorEnt = true;
										//								dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
									}
									if (bCrediPoint) bFecPrimPag = true;
									#endregion Fechas

									string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '{54}'", 
										(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value,
										(int)this.txtIdSucursal.Value, 
										(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
										(int)this.txtNumEstado.Value,
										bConsignacion,
										(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
										(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
										(int)this.cmbVendedor.Value,
										this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
										(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
										dtFecha.ToString("yyyyMMdd"), 
										this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), 
										(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion,
										this.txtNotas.Text.ToString(),
										bCrediPoint, dTotalCuotas, NumCuotas, (decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"),
										(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
										bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado, 
										dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, AutorizaEntrada, dValorCuota, (int)this.txtIdEF.Value,
										(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, this.txtNumeroCredito.Text.ToString());
									oCmdActualiza.CommandText = sSQL;
									this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
									#endregion Maestro

									#region Detalle
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
									{	
										int iPosicion = 0;
										if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicion = (int)dr.Cells["Posicion"].Value;

										string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}",
											(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value,
											(bool)dr.Cells["Manual"].Value, dr.Cells["SerieL"].Value.ToString(),
											(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
											(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
											(int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, iPosicion, (bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value);
										oCmdActualiza.CommandText = sSQLDetalle;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Detalle	

									string sSQLActDetalle = string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQLActDetalle;
									SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
									DataTable oDTDetalle = new DataTable();
									oDADetalle.Fill(oDTDetalle);
									this.ultraGrid1.DataSource = oDTDetalle;							
									#endregion Graba Factura					

									#region Seriales
									//							if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true))
									//							{
									if ((int)this.cmbBodega.Value != 56)
									{
										if (idTipoOrigen == 13 || idTipoOrigen == 11)
										{
											if (bNuevo)
											{
												#region Seriales
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
												{
													string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
													oCmdActualiza.CommandText = sSQLMovSerial;
													oCmdActualiza.ExecuteNonQuery();										
												}
												#endregion Seriales

												string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLActSeriales;
												SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
												DataTable oDTSeriales = new DataTable();
												oDASeriales.Fill(oDTSeriales);

												this.grdSerial.DataSource = oDTSeriales;
											}
										}
										else
										{
											if (!bConsignacion)
											{
												#region Grupos de seriales
												if (!bSerialesL)
												{
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
													{
														if (dr.Cells["idArticulo"].Value != DBNull.Value && (int) dr.Cells["idArticulo"].Value > 0)
														{
															int IdArticulo = (int) dr.Cells["idArticulo"].Value;
															int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;										
															int IPosicion = (int) dr.Cells["Posicion"].Value; 					

															foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
															{
																int IdArtSerial = (int) dr1.Cells["idArticulo"].Value;
																int IEstado = (int) dr1.Cells["Estado"].Value;
											
																if (IdArticulo == IdArtSerial && IPosicion == (int) dr1.Cells["Posicion"].Value)
																{
																	if (IEstado == 12)
																	{
																		int idSerial = 0;
																		int NIdSerial = 0;
																		if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;
																
																		if (idSerial < 1)
																		{
																			string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 12, {3}", idSerial, (int)dr1.Cells["idArticulo"].Value, dr1.Cells["Serial"].Value.ToString(), (int)this.cmbBodega.Value);
																			oCmdActualiza.CommandText = sSQLGSeriales;
																			NIdSerial = (int)oCmdActualiza.ExecuteScalar();
																		}
																		else NIdSerial = idSerial;

																		string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'LA SERIAL NO TIENE REGISTRO EN EL INVENTARIO.', 12, 0, {4}, {5}", NIdSerial, IdDetCompra, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
																		oCmdActualiza.CommandText = sSQLGDetalleSeriales;
																		oCmdActualiza.ExecuteNonQuery();

																		oCmdActualiza.CommandText = string.Format("Update Seriales Set Estado = {0} Where idSerial = {1}", 12, idSerial);
																		oCmdActualiza.ExecuteNonQuery();

																		string sSQLSM = string.Format("Insert Into SerialesManuales (idSerial, Serial, idArticulo, Bodega) Values ({0}, '{1}', {2}, {3})", 
																			NIdSerial, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value, (int)this.cmbBodega.Value);
																		oCmdActualiza.CommandText = sSQLSM;
																		oCmdActualiza.ExecuteNonQuery();	
																	}
																	else if (IEstado == 4 || IEstado == 2)
																	{
																		int idSerial = 0;
																		int idOrigen = 0;

																		if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
																		if (dr1.Cells["idOrigen"].Value != DBNull.Value) idOrigen = (int)dr1.Cells["idOrigen"].Value;

																		string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", idSerial, IdDetCompra, idOrigen, (int)this.cmbBodega.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
																		oCmdActualiza.CommandText = sSQLGDetalleSeriales;
																		oCmdActualiza.ExecuteNonQuery();
																
																		string sSQLActEstSerial = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", 4, idSerial, (int)this.cmbBodega.Value);
																		oCmdActualiza.CommandText = sSQLActEstSerial;																
																		oCmdActualiza.ExecuteNonQuery();

																		string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
																		oCmdActualiza.CommandText = sSQLConteo;
																		oCmdActualiza.ExecuteNonQuery();
																	}
																}
															}
														}
													}

													string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLActSeriales;
													SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
													DataTable oDTSeriales = new DataTable();
													oDASeriales.Fill(oDTSeriales);

													this.grdSerial.DataSource = oDTSeriales;											
												}
												#endregion Grupos de seriales
											}
										}
									}
									//							}
									#endregion Seriales

									if (idBloqueaTransacciones > 0)
									{
										string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
										oCmdActualiza.CommandText = sSQLBloqueo;
										oCmdActualiza.ExecuteNonQuery();
									}

									oTransaction.Commit();											
					
									#region Controles
									this.cmbFormaPago.Enabled = false;
									this.cmbBodega.Enabled = false;
									this.cmbVendedor.Enabled = false;
									this.cmbTipoRuc.Enabled = false;

									this.txtNumero.Enabled = false;
									this.txtRuc.Enabled = false;
									this.dtFecha.Enabled = false;
									this.txtNumeroCredito.Enabled = false;

									if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
									if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
									if (miAcceso.Editar) this.btnEditar.Enabled = true;									
									if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
									if (!miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;									
									this.btnGuardar.Enabled = false;
									this.btnCancelar.Enabled = true;
									this.btnClaveDescuento.Enabled = false;
									this.btnTemporal.Enabled = false;
									this.btnCliente.Enabled = false;

									this.ultraGrid1.DisplayLayout.Bands[0].Columns["SerieL"].CellActivation = Activation.AllowEdit;
									this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;
									this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.AllowEdit;
									this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;
					
									ModificaGrilla(false);
									HabilitaTab(false);

									bNuevo = false;
									bEdicion = false;
									bDescuentoActivo = false;
									idBloqueaTransacciones = 0;
									#endregion Controles
								}
								catch (Exception ex)
								{
									try
									{
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex2)
									{
										MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
									}
								}
								finally
								{
									oConexion.Close();
								}
							}
							#endregion Graba Factura
						}
					}
				}
			}
			#endregion Factura Temporal
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{			
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (bEdicion)
			{
				MessageBox.Show("No puede crear un asiento mientras edita la factura");				
			}
			else if (bAnulado) 
			{
				MessageBox.Show("No puede crear un asiento de una transacción Anulada");				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Crear Asiento de esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}			
			else
			{					
				string stExisteAsiento = string.Format("Exec AsientoMaestroExiste {0}", (int)this.txtNumIdCompra.Value);
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				}
				else
				{
					stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 2", true);
					if (stMensaje.Length > 0)
					{
						if (DialogResult.Yes == MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{						
							stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 3", true);

							#region Crea Asiento
							this.Cursor = Cursors.WaitCursor;
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", (int)this.txtNumIdCompra.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

							MessageBox.Show("Asiento Generado con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
							this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
							Cursor = Cursors.Default;							
							#endregion Crea Asiento
						}
					}				
					else if (DialogResult.Yes == MessageBox.Show("¿Desea crear Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{	
						#region Crea Asiento
						this.Cursor = Cursors.WaitCursor;
						string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", (int)this.txtNumIdCompra.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

						MessageBox.Show("Asiento Generado con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
						this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
						Cursor = Cursors.Default;
						#endregion Crea Asiento
					}
				}
			}
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (miAcceso.CrearAsiento)
				{
					if ((int) this.txtNumIdAsiento.Value == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					else
					{					
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 0) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdAsiento.Value), true))
						{
							Cursor = Cursors.WaitCursor;
							frmAsientos Asientos = new frmAsientos((int)this.txtNumIdAsiento.Value);
							Asientos.MdiParent = MdiParent;
							Asientos.Show();
							Cursor = Cursors.Default;
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Factura esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Factura fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
					}
				}
			}
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsLetterOrDigit(e.KeyChar))
			{
				e.Handled = false;
			}			
			else
			{
				e.Handled = true;
			}
		}

		private void txtSerial_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void PruebaFirma()
		{
			CspParameters cspParams = new CspParameters();
			cspParams.KeyContainerName = @"C:/santiago_fabian_tejada_rosero.p12";

			RSACryptoServiceProvider rsaKey = new RSACryptoServiceProvider(cspParams);

			XmlDocument xmlDoc = new XmlDocument();
			
			xmlDoc.PreserveWhitespace = true;
			xmlDoc.Load(@"C:/Factprueba.xml");

			SignedXml signedXml = new SignedXml(xmlDoc);

			signedXml.SigningKey = rsaKey;
			
			Reference reference = new Reference();			
			reference.Uri = @"C:\Factprueba.xml";

			XmlDsigEnvelopedSignatureTransform env = new XmlDsigEnvelopedSignatureTransform();
			reference.AddTransform(env);

			signedXml.AddReference(reference);

			KeyInfo keyInfo = new KeyInfo();
			keyInfo.AddClause(new RSAKeyValue((RSA)rsaKey));	
			signedXml.KeyInfo = keyInfo;
			
			signedXml.ComputeSignature();

			XmlElement xmlDigitalSignature = signedXml.GetXml();

			xmlDoc.DocumentElement.AppendChild(xmlDoc.ImportNode(xmlDigitalSignature, true));

			xmlDoc.Save(@"C:/Factprueba.xml");
		}

		private void btnFE_Click(object sender, System.EventArgs e)
		{		
			#region Variables
			int iAmbiente = 0;
			int iTipoEmision = 0;
			string sRazonSocial = "";
			string sNombreComercial = "";
			string sRUC = "";
			string sClaveAcceso = "";
			string sCodDoc = "";
			string sEstab = "";
			string sPtoEmi = "";
			string sDirMatriz = "";
			string sDirEstablecimiento = "";
			int iResolucion = 0;
			string sCodDocClte = "";
			decimal dSubtotal = (decimal)this.txtIva0.Value + (decimal)this.txtIva.Value;
			decimal dDescuento = (decimal)this.txtDescIva0Total.Value + (decimal)this.txtDescIvaTotal.Value;
			decimal dIVA0 = (decimal)this.txtIva0.Value;
			decimal dIVA = (decimal)this.txtIva.Value;
			decimal dIVATotal = (decimal)this.txtIvaTotal.Value;
			decimal dTotal = (decimal)this.txtTotal.Value;
			#endregion Variables
						
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			
			#region Consulta Informacion Tributaria
			string sSQL = string.Format("Exec FEInformacionTributaria {0}, {1}, {2}, '{3}', '{4}', {5}", 1, 1, (int)this.cmbBodega.Value, dtFecha.ToString("ddMMyyyy"), this.txtNumero.Text.ToString(), (int)this.txtNumIdCliente.Value);
			SqlDataReader drInfoTributaria = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drInfoTributaria.Read();
			if (drInfoTributaria.HasRows)
			{				
				iAmbiente = (int)drInfoTributaria.GetValue(0);
				iTipoEmision = (int)drInfoTributaria.GetValue(1);
				sRazonSocial = drInfoTributaria.GetValue(2).ToString();
				sNombreComercial = drInfoTributaria.GetValue(3).ToString();
				sRUC = drInfoTributaria.GetValue(4).ToString();
				sClaveAcceso = drInfoTributaria.GetValue(5).ToString();
				sCodDoc = drInfoTributaria.GetValue(6).ToString();
				sEstab = drInfoTributaria.GetValue(7).ToString();
				sPtoEmi = drInfoTributaria.GetValue(8).ToString();
				sDirMatriz = drInfoTributaria.GetValue(9).ToString();
				sDirEstablecimiento = drInfoTributaria.GetValue(10).ToString();
				iResolucion = (int)drInfoTributaria.GetValue(11);
				sCodDocClte = drInfoTributaria.GetValue(12).ToString();
			}
			drInfoTributaria.Close();
			#endregion Consulta Informacion Tributaria

			#region XML
			#region Encabezado
			string stArchivo = @"C:/Factprueba.xml";
			string stArchivo1 = @"C:/Factprueba1.xml";
			XmlTextWriter XMLFactura = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));
//			XMLFactura.Formatting = Formatting.Indented;
			XMLFactura.WriteStartDocument();
			XMLFactura.WriteStartElement("factura");
			XMLFactura.WriteAttributeString("id", "comprobante");
      XMLFactura.WriteAttributeString("version", "1.0.0");
			#endregion Encabezado
			
			#region Informacion Tributaria
			XMLFactura.WriteStartElement("infoTributaria");

			XMLFactura.WriteElementString("ambiente", iAmbiente.ToString());
			XMLFactura.WriteElementString("tipoEmision", iTipoEmision.ToString());
			XMLFactura.WriteElementString("razonSocial", sRazonSocial);
			XMLFactura.WriteElementString("nombreComercial", sNombreComercial);
			XMLFactura.WriteElementString("ruc", sRUC);
			XMLFactura.WriteElementString("claveAcceso", sClaveAcceso);
			XMLFactura.WriteElementString("codDoc", sCodDoc);
			XMLFactura.WriteElementString("estab", sEstab);
			XMLFactura.WriteElementString("ptoEmi", sPtoEmi);
			XMLFactura.WriteElementString("secuencial", this.txtNumero.Text.ToString());
			XMLFactura.WriteElementString("dirMatriz", sDirMatriz);

			XMLFactura.WriteEndElement();
			#endregion Informacion Tributaria

			#region Informacion Factura
			XMLFactura.WriteStartElement("infoFactura");

			XMLFactura.WriteElementString("fechaEmision", dtFecha.ToString("dd/MM/yyyy"));
			XMLFactura.WriteElementString("dirEstablecimiento", sDirEstablecimiento);
			XMLFactura.WriteElementString("contribuyenteEspecial", iResolucion.ToString());
			XMLFactura.WriteElementString("obligadoContabilidad", "SI");
			XMLFactura.WriteElementString("tipoIdentificacionComprador", sCodDocClte);
			XMLFactura.WriteElementString("razonSocialComprador", this.txtNombre.Text.ToString());
			XMLFactura.WriteElementString("identificacionComprador", this.txtRuc.Text.ToString());
			XMLFactura.WriteElementString("totalSinImpuestos", dSubtotal.ToString("0.00", us));
			XMLFactura.WriteElementString("totalDescuento", dDescuento.ToString("0.00", us));

			#region Impuestos
			XMLFactura.WriteStartElement("totalConImpuestos");

			if ((decimal)this.txtIva0.Value > 0)
			{
				XMLFactura.WriteStartElement("totalImpuesto");

				XMLFactura.WriteElementString("codigo", "2");
				XMLFactura.WriteElementString("codigoPorcentaje", "0");
				XMLFactura.WriteElementString("baseImponible", dIVA0.ToString("0.00", us));
				XMLFactura.WriteElementString("valor", "0.00");

				XMLFactura.WriteEndElement();
			}

			if ((decimal)this.txtIva.Value > 0)
			{
				XMLFactura.WriteStartElement("totalImpuesto");

				XMLFactura.WriteElementString("codigo", "2");
				XMLFactura.WriteElementString("codigoPorcentaje", "2");
				XMLFactura.WriteElementString("baseImponible", dIVA.ToString("0.00", us));
				XMLFactura.WriteElementString("valor", dIVATotal.ToString("0.00", us));

				XMLFactura.WriteEndElement();
			}

			XMLFactura.WriteEndElement();
			#endregion Impuestos
      
			XMLFactura.WriteElementString("propina", "0.00");
			XMLFactura.WriteElementString("importeTotal", dTotal.ToString("0.00", us));
			XMLFactura.WriteElementString("moneda", "DOLAR");

			XMLFactura.WriteEndElement();
			#endregion Informacion Factura

			#region Detalles
			XMLFactura.WriteStartElement("Detalles");
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				XMLFactura.WriteStartElement("Detalle");
	
				decimal dCantidad = (decimal)dr.Cells["Cantidad"].Value;
				decimal dPrecio = (decimal)dr.Cells["Precio"].Value;
				decimal dDescuentoUnit = (decimal)dr.Cells["DescuentoArt"].Value;
				decimal dPrecioTotal = (decimal)dr.Cells["Subtotal"].Value - (decimal)dr.Cells["DescuentoArt"].Value;

				XMLFactura.WriteElementString("codigoPrincipal", dr.Cells["Codigo"].Value.ToString());
				XMLFactura.WriteElementString("descripcion", dr.Cells["Articulo"].Value.ToString());
				XMLFactura.WriteElementString("cantidad", dCantidad.ToString("0.00", us));
				XMLFactura.WriteElementString("precioUnitario", dPrecio.ToString("0.00", us));
				XMLFactura.WriteElementString("descuento", dDescuentoUnit.ToString("0.00", us));
				XMLFactura.WriteElementString("precioTotalSinImpuesto", dPrecioTotal.ToString("0.00", us));

				#region Impuestos
				XMLFactura.WriteStartElement("Impuestos");

				if ((decimal)dr.Cells["Impuesto"].Value == 0)
				{
					XMLFactura.WriteStartElement("totalImpuesto");

					decimal dDescArt = (decimal) dr.Cells["DescuentoArt"].Value;
					decimal dSubtArt = (decimal) dr.Cells["Subtotal"].Value;
					decimal dSubtLineal = dSubtArt - dDescArt;

					XMLFactura.WriteElementString("codigo", "2");
					XMLFactura.WriteElementString("codigoPorcentaje", "0");
					XMLFactura.WriteElementString("tarifa", "0");
					XMLFactura.WriteElementString("baseImponible", dSubtLineal.ToString("0.00", us));
					XMLFactura.WriteElementString("valor", "0.00");

					XMLFactura.WriteEndElement();
				}

				if ((decimal)dr.Cells["Impuesto"].Value > 0)
				{
					XMLFactura.WriteStartElement("totalImpuesto");

					decimal dDescArt = (decimal) dr.Cells["DescuentoArt"].Value;
					decimal dSubtArt = (decimal) dr.Cells["Subtotal"].Value;
					decimal dSubtLineal = dSubtArt - dDescArt;
					decimal dIvaLineal = dSubtLineal * 12 /100;

					XMLFactura.WriteElementString("codigo", "2");
					XMLFactura.WriteElementString("codigoPorcentaje", "2");
					XMLFactura.WriteElementString("tarifa", "12");
					XMLFactura.WriteElementString("baseImponible", dSubtLineal.ToString("0.00", us));
					XMLFactura.WriteElementString("valor", dIvaLineal.ToString("0.00", us));

					XMLFactura.WriteEndElement();
				}

				XMLFactura.WriteEndElement();
				#endregion Impuestos

				XMLFactura.WriteEndElement();
			}

			XMLFactura.WriteEndElement();
			#endregion Detalles

			XMLFactura.WriteEndElement(); 
			XMLFactura.WriteEndDocument();

			XMLFactura.Flush();
			XMLFactura.Close();
			#endregion XML

			PruebaFirma();

			#region Firma	
			/*
			RSACryptoServiceProvider Clave = new RSACryptoServiceProvider();
			//FirmaXML(stArchivo, stArchivo1, Clave);
			//bool result = VerificaFirma(stArchivo1);

			XmlDocument documentoXML = new XmlDocument();
			documentoXML.PreserveWhitespace = true;
			documentoXML.Load(@"C:/Factprueba.xml");

			CspParameters cspParams = new CspParameters();
			cspParams.KeyContainerName = "santiago_fabian_tejada_rosero.p12";			
			RSACryptoServiceProvider Key = new RSACryptoServiceProvider(cspParams);

			SignedXml xmlFirmado = new SignedXml(documentoXML);
 
			// Agregue la clave RSA de firma al objeto SignedXml.
			xmlFirmado.SigningKey = Key;
 
			// Cree un objeto Reference que describa qué se debe firmar.
			// Para firmar el documento completo, establezca la propiedad Uri como "".
			Reference referencia = new Reference();
			referencia.Uri = "";
 
			// Agregue un objeto XmlDsigEnvelopedSignatureTransform al objeto Reference.
			// Una transformación permite al comprobador representar los datos XML
			// de idéntico modo que el firmante. Los datos XML se pueden representar de distintas maneras,
			// por lo que este paso es vital para la comprobación.
			XmlDsigEnvelopedSignatureTransform transformacionENV = new XmlDsigEnvelopedSignatureTransform();
			referencia.AddTransform(transformacionENV);
 
			// Agregue el objeto Reference al objetoSignedXml.
			xmlFirmado.AddReference(referencia);
 
			// Llame al método ComputeSignature para calcular la firma.
			xmlFirmado.ComputeSignature();
 
			// Recupere la representación XML de la firma (un elemento <Signature>)
			// y guárdela en un nuevo objeto XmlElement.
			XmlElement firmaDigitalXML = xmlFirmado.GetXml();
 
			// Anexe el elemento al objeto XmlDocument.
			documentoXML.DocumentElement.AppendChild(documentoXML.ImportNode(firmaDigitalXML, true));
			firmaDigitalXML.WriteTo(XMLFactura);
			//SignedXml xmlFirmado = new SignedXml(xmlDoc);
 
			// Busque el elemento <signature> y cree un nuevo objeto XmlNodeList.
			XmlNodeList listaNodos = documentoXML.GetElementsByTagName("Signature");
 			
			// Cargue el XML del primer elemento <signature> en el objeto SignedXml.
			xmlFirmado.LoadXml((XmlElement)listaNodos[0]);

			xmlFirmado.CheckSignature(Key);
			*/
			#endregion Firma			
		}

		public static void FirmaXML(string FileName, string SignedFileName, RSA Key)
		{
			// Create a new XML document.
			XmlDocument doc = new XmlDocument();

			// Format the document to ignore white spaces.
			doc.PreserveWhitespace = false;

			// Load the passed XML file using it's name.
			doc.Load(new XmlTextReader(FileName));

			// Create a SignedXml object.
			SignedXml signedXml = new SignedXml(doc);

			// Add the key to the SignedXml document. 
			signedXml.SigningKey = Key;

			// Create a reference to be signed.
			Reference reference = new Reference();
			reference.Uri = "";

			// Add an enveloped transformation to the reference.
			XmlDsigEnvelopedSignatureTransform env = new XmlDsigEnvelopedSignatureTransform();
			reference.AddTransform(env);

			// Add the reference to the SignedXml object.
			signedXml.AddReference(reference);

		
			// Add an RSAKeyValue KeyInfo (optional; helps recipient find key to validate).
			KeyInfo keyInfo = new KeyInfo();
			keyInfo.AddClause(new RSAKeyValue((RSA)Key));
			signedXml.KeyInfo = keyInfo;

			// Compute the signature.
			signedXml.ComputeSignature();

			// Get the XML representation of the signature and save
			// it to an XmlElement object.
			XmlElement xmlDigitalSignature = signedXml.GetXml();

			// Append the element to the XML document.
			doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, true));
		
		
			if (doc.FirstChild is XmlDeclaration)  
			{
				doc.RemoveChild(doc.FirstChild);
			}

			// Save the signed XML document to a file specified
			// using the passed string.
			XmlTextWriter xmltw = new XmlTextWriter(SignedFileName, new UTF8Encoding(false));
			doc.WriteTo(xmltw);
			xmltw.Close();
		}

		public static Boolean VerificaFirma(string Name)
		{
			// Create a new XML document.
			XmlDocument xmlDocument = new XmlDocument();

			// Format using white spaces.
			xmlDocument.PreserveWhitespace = true;

			// Load the passed XML file into the document. 
			xmlDocument.Load(Name);

			// Create a new SignedXml object and pass it
			// the XML document class.
			SignedXml signedXml = new SignedXml(xmlDocument);

			// Find the "Signature" node and create a new
			// XmlNodeList object.
			XmlNodeList nodeList = xmlDocument.GetElementsByTagName("Signature");

			// Load the signature node.
			signedXml.LoadXml((XmlElement)nodeList[0]);

			// Check the signature and return the result.
			return signedXml.CheckSignature();
		}

		private void txtSerieRetencion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumeroRetencion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkEntregado_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && !bNumeracionAutomatica)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;					
				}
				else
				{
					char Pad = '0';
					this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);

					string sSQL = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)this.txtNumIdCompra.Value, this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
					string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL, true);

					if (sMensaje.ToString().Length > 0)
					{
						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumero.Focus();
						e.Cancel = true;				
					}
				}
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) 
			{
				if (this.txtRuc.Enabled) this.txtRuc.Focus();
				else if (miAcceso.CambiarFecha) this.dtFecha.Focus();
				else this.btnGuardar.Focus();
			}
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (!bNumeracionAutomatica)
				{
					char Pad = '0';
					this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);
				}
			}
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
			{
				int IdSerial = 0;
				int IdDetalleSerial = 0;
				bool SerialRepetida = false;
				int CantSerie = 0;

				#region Cuenta si hay articulos en el grid
				if (this.ultraGrid1.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese los artículos para cargar seriales.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Focus();
					return;
				}
				#endregion Cuenta si hay articulos en el grid

				#region Valida Selección 
				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Focus();
					return;
				}
				#endregion Valida Selección 

				int IdArticulo = (int) this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
				int IPosicion = (int) this.ultraGrid1.ActiveRow.Cells["Posicion"].Value; 

				#region Valida serial repetida
				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
				//				{
				//					string Serial = (string) dr.Cells["Serie"].Value;
				//					if (Serial.Trim() == this.txtSerial.Text.Trim()) SerialRepetida = true;
				//				}
				//				
				//				if (SerialRepetida == true)
				//				{
				//					MessageBox.Show("La serial ingresada ya esta registrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//					this.txtSerial.Text = "";
				//					return;
				//				}
				#endregion Valida serial repetida
				
				#region Valida selección de articulo en el detalle
				if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idArticulo"].Value == DBNull.Value)
				{
					MessageBox.Show("Seleccione el Artículo", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
				#endregion Valida selección de articulo en el detalle

				#region Valida cantidad seriales
				if (SerialRepetida == false)
				{
					if (!ModSerial)
					{
						int CantArticulo = 0;
						CantArticulo = (int)this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value;
												      
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
						{
							int idArtSerial = (int)dr1.Cells["idArticulo"].Value;
							int iPosSerial = (int)dr1.Cells["Posicion"].Value;

							if (IdArticulo == idArtSerial && IPosicion == iPosSerial)
							{
								CantSerie++;
							}
						}

						if (CantSerie >= CantArticulo)
						{
							MessageBox.Show("Ya ha ingresado las seriales para este articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.txtSerial.Text = "";
							this.txtSerial.Focus();
							return;
						}
					}
				}
				#endregion Valida cantidad seriales

				#region Valida serial existente en el grid
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdSerial.Rows.All)
				{
					string sSerial = (string)dr.Cells["Serial"].Value;
					if ((string)this.txtSerial.Text == (string)sSerial)
					{
						MessageBox.Show("Ya registro esta serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtSerial.Text = "";
						this.txtSerial.Focus();
						return;
					}
				}
				#endregion Valida serial existente en el grid

				#region Valida si la Serial Existe en el Sistema o no Pertenece al Articulo Seleccionado
				string sSQLVSA = string.Format("Exec ValidaSerialArticulo '{0}', {1}", this.txtSerial.Text.ToString(), IdArticulo);
				string sMensajeVSA = Funcion.sEscalarSQL(cdsSeteoF, sSQLVSA);

				if (sMensajeVSA.Length > 0)
				{
					MessageBox.Show(sMensajeVSA, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
				#endregion Valida si la Serial Existe en el Sistema o no Pertenece al Articulo Seleccionado

				#region Valida Serial existente en la base de datos se bloqueo solo para e/r mercaderia
//				int Cont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValExistenciaSerial '{0}', {1}, {2}, {3}", (string)this.txtSerial.Text, IdArticulo, (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value));
//				if(Cont > 0) 
//				{
//					MessageBox.Show("La Serial ya esta registrada en el sistema.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//					this.txtSerial.Text = "";
//					this.txtSerial.Focus();
//					return;
//				}
				#endregion Valida Serial existente en la base de datos se bloqueo solo para e/r mercaderia
				
				#region Valida si la Serial esta en Orden de Compra - En Transito - Pedidos - Reservas - Consignaciones - Ventas 
				int iCont = 0;
				string sMens = "";
				string sSQLValSerEsp = string.Format("Exec ValidaSerialOrdCmpEnTransito '{0}', {1}, {2}, 0", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega.Value);
				SqlDataReader drV = Funcion.rEscalarSQL(cdsSeteoF, sSQLValSerEsp, true);
				drV.Read();
				if (drV.HasRows)
				{
					iCont = drV.GetInt32(0);
					sMens = drV.GetString(1);
				}
				drV.Close();
								
				if (iCont > 0)
				{
					MessageBox.Show(sMens, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerie.Focus();
					return;
				}
				#endregion Valida si la Serial esta en Orden de Compra - En Transito - Pedidos - Reservas - Consignaciones - Ventas

				#region Valida Serial En Bodega
//				string sSQLVSEB = string.Format("Exec ValidaSerialEnBodega '{0}', {1}, {2}", this.txtSerial.Text.ToString(), (int)this.cmbBodega.Value, IdArticulo);
//				string sMensajeVSEB = Funcion.sEscalarSQL(cdsSeteoF, sSQLVSEB);
//
//				if (sMensajeVSEB.Length > 0)
//				{
//					MessageBox.Show(sMensajeVSEB, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.txtSerial.Text = "";
//					this.txtSerial.Focus();
//					return;
//				}	
				#endregion Valida Serial En Bodega

				#region Carga Grid de Seriales
				string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}", (string) this.txtSerial.Text, IdArticulo, (int)this.cmbBodega.Value);
				int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
				if (iCtaSerial == 1)
				{	
					IdSerial = 0;
					IdDetalleSerial = 0;
					string sSQLIdSerial = string.Format("Exec BuscaIdSerial '{0}', {1}, {2}", (string) this.txtSerial.Text, IdArticulo, (int)this.cmbBodega.Value);
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						IdSerial = drIdSerial.GetInt32(0);
						IdDetalleSerial = drIdSerial.GetInt32(1);
					}
					drIdSerial.Close();
					
					if (!ModSerial) 
					{																		
						this.grdSerial.Rows.Band.AddNew();
						this.grdSerial.ActiveRow.Cells["idSerial"].Value = IdSerial;
						this.grdSerial.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.grdSerial.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
						this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
						this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
					}
					else
					{
						this.grdSerial.ActiveRow.Cells["idSerial"].Value = IdSerial;						
						this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;						
						this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
						this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
						this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
					}
				}
				else
				{
//					string sSQLVSM = string.Format("Exec ValidacionDeSerialSinMovimiento '{0}', {1}", this.txtSerial.Text.ToString(), IdArticulo);
//					int iValRes = Funcion.iEscalarSQL(cdsSeteoF, sSQLVSM);
//					if (iValRes > - 1)
//					{
//						if (!ModSerial) 
//						{	
//							this.grdSerial.Rows.Band.AddNew();
//							this.grdSerial.ActiveRow.Cells["idSerial"].Value = iValRes;
//							this.grdSerial.ActiveRow.Cells["idDetalleSerial"].Value = -1;
//							this.grdSerial.ActiveRow.Cells["idOrigen"].Value = -1;
//							this.grdSerial.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
//							this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();
//							this.grdSerial.ActiveRow.Cells["Estado"].Value = 12;	
//							this.txtNumEstado.Value = 11;
//							this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
//						}
//						else
//						{
//							this.grdSerial.ActiveRow.Cells["idSerial"].Value = iValRes;
//							this.grdSerial.ActiveRow.Cells["idOrigen"].Value = -1;
//							this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();
//							this.grdSerial.ActiveRow.Cells["Estado"].Value = 12;
//							this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
//						}
//
//						this.lblEstado.Text = "CONTABILIZADO";
//					}
//					else
//					{
						MessageBox.Show("La serial no existe, Esta en otro Local o no pertenece al artículo seleccionado,\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtSerial.Focus();
						return;
//					}			
				}
			
				this.txtSerial.Text = "";
				#endregion Carga Grid de Seriales

				this.txtSerial.Text = "";

				this.txtSerial.Focus();

				ModSerial = false;
			}
		}

		private void frmVenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtNumIdCompra.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (bImpreso) 
																	 {
																		 if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																	 }
																	 else if (miAcceso.ReImprimir) if (this.btnImprimir.Enabled) this.btnImprimir.Enabled = true;
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0703");

						this.Consultar((int)this.txtNumIdCompra.Value);					
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);
				if (e.Control && e.KeyCode == Keys.S) 
				{
					this.utcTransaccion.Tabs[0].Selected = true;
					this.txtSerial.Focus();
				}
				if (e.Control && e.KeyCode == Keys.R) 
				{
					this.utcTransaccion.Tabs[1].Selected = true;
					this.txtAutorizacionRetencion.Focus();
				}
				if (e.Control && e.KeyCode == Keys.N)
				{
					this.utcTransaccion.Tabs[4].Selected = true;
					this.txtNotas.Focus();
				}
			}
		}

		private void grdSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (grdSerial.ActiveCell == null) return;
			try
			{				
				switch(e.KeyValue)
				{								
						#region Flechas
					case (int) Keys.Up:
						if (grdSerial.ActiveCell.DroppedDown) return;

						grdSerial.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdSerial.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grdSerial.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grdSerial.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdSerial.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grdSerial.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grdSerial.ActiveCell.DroppedDown) return;			
						grdSerial.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdSerial.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grdSerial.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						grdSerial.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdSerial.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grdSerial.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grdSerial.ActiveCell.DroppedDown == false) grdSerial.ActiveCell.DroppedDown = true;
						//						grdSerial.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
			}
		}

		private void grdSerial_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void grdSerial_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				for (int i=0; i<e.Rows.Length; i++)
				{	
					if ((int)e.Rows[i].Cells["idDetalleSerial"].Value > 0)
					{
						string sSQL = string.Format("Exec ValSerialBorrar {0}, 1", (int)e.Rows[i].Cells["idSerial"].Value);
						bool bValida = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);

						if (!bValida)
						{
							MessageBox.Show(string.Format("No puede Eliminar esta serial '{0}' ya ha sido utilizada", e.Rows[i].Cells["Serial"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
							e.Cancel = true;
							return;
						}
					}

					if (DialogResult.No ==	MessageBox.Show(string.Format("¿Esta seguro de borrar la serial '{0}'?", e.Rows[i].Cells["Serial"].Value.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;
					else
					{
						if ((int)e.Rows[i].Cells["idDetalleSerial"].Value > 0)
						{
							string sSQL = string.Format("Exec BorrarSeriales 1, {0}, {1}, {2}, {3}, {4}", 
								(int)e.Rows[i].Cells["idSerial"].Value, (int)e.Rows[i].Cells["idDetalleSerial"].Value, (int)e.Rows[i].Cells["Estado"].Value, (int)this.txtNumIdCompra.Value, (int)e.Rows[i].Cells["idDetCompra"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					
						e.DisplayPromptMsg = false;
					}
				}
			}
		}

		private void frmVenta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo" || e.Cell.Column.ToString() == "Articulo")
			{
				if (!bModificaColor && (bool)e.Cell.Row.Cells["Principal"].Value)
				{
					MessageBox.Show("No puede modificar el Producto Principal de un Combo o Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}

				if (!bModificaColor && (int)e.Cell.Row.Cells["Tipo"].Value > 0 && !miAcceso.BModArtCmbProm)
				{
					MessageBox.Show("No puede modificar Productos de un Combo o Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}

//			if (e.Cell.Column.ToString() == "DescuentoPorc")
//			{
//				if (!bDescuentoActivo)  
//				{
//					MessageBox.Show("No Esta Autorizado a Modificar Descuento,\n\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					e.Cancel = true;
//				}
//			}

//			if (e.Cell.Column.ToString() == "DescuentoPorc")
//			{
//				if ((int)e.Cell.Row.Cells["idCombo"].Value > 0)  
//				{
//					if (!bDescuentoActivo)  
//					{
//						MessageBox.Show("Producto Promocional No puede Aplicar otro Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						e.Cancel = true;
//					}
//				}				
//			}
		}

		bool bModificaColor = false;

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbFormaPago.Value != 9)
				{
					if ((int)this.ultraGrid1.ActiveRow.Cells["Tipo"].Value > 0)
					{
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaColoresArticulosPromocion {0}", (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value)) <= 1)
						{
							MessageBox.Show(string.Format("El Articulo {0} solo tiene un color", this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
						}
						else
						{
							string sSQL = string.Format("Exec ListaArticulosPromocionPorColor {0}, {1}", (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value, (int)this.cmbBodega.Value);
							if (Funcion.dvProcedimiento(cdsSeteoF, sSQL).Count == 0)
							{
								MessageBox.Show("Este Articulo no tiene Existencia en otros Colores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
							}
							else
							{
								using (frmColoresProductos Colores = new frmColoresProductos((int)ultraGrid1.ActiveRow.Cells["idArticulo"].Value, (int)this.cmbBodega.Value))
								{				
									if (DialogResult.OK == Colores.ShowDialog())
									{
										bModificaColor = true;
										ultraGrid1.ActiveRow.Cells["idArticulo"].Value = (int)Colores.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
										ultraGrid1.ActiveRow.Cells["Codigo"].Value = Colores.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
										ultraGrid1.ActiveRow.Cells["Articulo"].Value = Colores.ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString();
										bModificaColor = false;
									}
								}
							}
						}
					}
				}
			}
		}

		private void dtFechaEntrega_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtpHoraEntrega_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbDirecciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnImprimeDctos_Click(object sender, System.EventArgs e)
		{			
			decimal dValor = (decimal)this.txtTotal.Value - (decimal)this.txtCuotaDeEntrada.Value;
			string Fecha = dtFechPrimPago.ToString();//.ToString("yyyyMMdd");			
			DocsPrefactura Dctos = new DocsPrefactura ((int)this.txtNumIdCompra.Value, dValor, NumCuotas, (int)this.txtNumIdCliente.Value, Fecha, (DateTime)this.dtFecha.Value, (int)this.txtIdEF.Value, this.txtRuc.Text.ToString());
			Dctos.ShowDialog();		
		}

		private void btnClaveDescuento_Click(object sender, System.EventArgs e)
		{
			using (frmVentaClaveDescuento Clave = new frmVentaClaveDescuento())
			{				
				if (DialogResult.OK == Clave.ShowDialog())
				{
					bDescuentoActivo = true;
				}
			}
		}

		private void dtFechaRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnRetener.Focus();
		}

		private void dtFechaRetencion_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFechaRetencion_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) if (this.dtFechaRetencion.Value == System.DBNull.Value) this.dtFechaRetencion.Value = (DateTime)this.dtFecha.Value;
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.dtFecha.Value != System.DBNull.Value)
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;

					this.dtFechaRetencion.Value = (DateTime)this.dtFecha.Value;
					this.dtFechaRetencion.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
					this.dtFechaRetencion.CalendarInfo.MaxDate = dtFecha.AddDays(5);
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la fecha del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();			
					e.Cancel = true;
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;
				}
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) if (this.dtFecha.Value != System.DBNull.Value) this.dtFechaEntrega.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
		}

		private void cmbCaducidad_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}

