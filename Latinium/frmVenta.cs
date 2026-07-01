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
using System.Threading;
using Latinium.Services;
using Latinium.Models;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmVenta.
	/// </summary>
	public class frmVenta : DevExpress.XtraEditors.XtraForm
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
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen;
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
		string CodigoEntrega ="";
		bool bEdicion = false;
		decimal dAnticipo = 0;
		bool bUsoAnticipo = false;
		bool bPromocionCombo = false;
		bool bPromocion = false;
		bool bCombo = false;
		int IdSigno = 2;
		int idBloqueaTransacciones = 0;
		int idBloqueaTransacciones1 = 0;
		bool bSerialesL = false;
		bool bFacturacManual = false;
		bool bSerialNoExiste = false;
		bool bConsignacion = false;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private System.Windows.Forms.Label lblEstado;
		bool bIncluyeIva = false;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnAnular;
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
		private System.Windows.Forms.Button btnCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdproyectoZoom;
		private System.Windows.Forms.Button btnTemporal;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.TextBox txtSerial;
		private System.Windows.Forms.GroupBox gbBotones;
		bool bAnulado = false;
		int idFe_EstadoDocumento = 0;
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
		bool bCuponWeb = false;
		decimal dPorcentajeCuponWeb = 0.00m;
		int idCuponWeb = 0;
		bool bClienteFacturaBajoCosto = false;
		int PointMovil = 0;
		#endregion Variables

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
		bool bDescuentoActivo = false;
		decimal dClavePorcDscto = 0.00m;
		bool bVeriricador = false;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirFactura;
		private System.Windows.Forms.MenuItem mnuDocsCrediPoint;
		private System.Windows.Forms.MenuItem mnuImprimirGarantiaExtendida;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpHoraEntrega;
		decimal dPorcIva = 0.00m;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCourier;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumGuia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccionEntrega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorEntrega;
		private System.Windows.Forms.Label label13;
		private DevExpress.XtraEditors.RadioGroup optTipoVenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		bool bValidaTipoVenta = false;
		private System.Windows.Forms.ContextMenu mnuBI;
		private System.Windows.Forms.MenuItem mnuBIPedidos;
		private System.Windows.Forms.MenuItem mnuBICotizaciones;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClienteNombre;
		bool bDescuentoLineal = false;
		int iNumCompras = 0;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega1;
		private System.Windows.Forms.Label lblresponsablecourier;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label Pieza;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntregaCuriermbo1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtresponsablecurier;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtphoracurier;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox txtNumTransaccion;
		private System.Windows.Forms.Button btnCliente2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente1;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Button btnImprimeED;
		private System.Windows.Forms.Button btnValidarMail;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProteccionDD;
		private System.Windows.Forms.MenuItem mnuImprimirDYD;
		private System.Windows.Forms.MenuItem mnuImprimirRIDE;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.MenuItem mnuDocsCrediPointBlanco;
		private System.Windows.Forms.MenuItem mnuImprimirCP;
		private System.Windows.Forms.MenuItem mnuImprimirPR;
		private System.Windows.Forms.MenuItem mnuImprimirASD;
		private System.Windows.Forms.MenuItem mnuImprimirGEP;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem mnuBILaboratorio;
		private System.Windows.Forms.MenuItem mnuImprimirContrato;
		private System.Windows.Forms.MenuItem mnuImprimirNuovo;
		int iDiasPosteriores = 0;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFreelance;
		private System.Windows.Forms.Button btnSubir;
		bool bFundas=false;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregaLocal;
		private System.Windows.Forms.Button btnEntrega;
		private System.Windows.Forms.Label lblRecurrente;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNumContadoCredito;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFechaPrimerPago;
		private System.Windows.Forms.Label lblYupana;
		private int idPersonalGlobal = 0;


		public frmVenta(int IdCompraV)
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

		bool bAutoConsumo = false;

		public frmVenta(int IdCompraV, bool BAutoConsumo)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idCompraV = IdCompraV;
			bAutoConsumo = BAutoConsumo;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmVenta()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int iBodegaFacDT = 0;
		string sRuc = "";
		int idTecnico = 0;
		bool bServicioTecnico = false;
		string sCodigoServTecn = "";
		decimal dPrecio = 0.00m;
		int iEstadoDT = 0;
		int idFormaPagoST = 0;
		int iBodegaDestino = 0;
		int iAnticipoCDP = 0;
		int iAsientoCDP = 0;
		public static string sRideFactura = "RIDEFactura.rpt";

		public frmVenta(string SRuc, int IBodegaFacDT, int IdTecnico, bool BServicioTecnico, string SCodigoServTecn, decimal DPrecio, int IEstadoDT, int IdFormaPagoST)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			iBodegaFacDT = IBodegaFacDT;
			sRuc = SRuc;
			idTecnico = IdTecnico;
			bServicioTecnico = BServicioTecnico;
			sCodigoServTecn = SCodigoServTecn;
			dPrecio = DPrecio;
			iEstadoDT = IEstadoDT;
			idFormaPagoST = IdFormaPagoST;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/* PROVIENE DE SOLICITUD DE NOTAS DE CREDITO */
		int idCompraOrigen = 0;
		int idMotivo = 0;
		decimal dTotalOrigen = 0.00m;
		
		public frmVenta(int IdCompraOrigen, bool bNotaCredito, int IdMotivo, decimal DTotalOrigen)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			idCompraOrigen = IdCompraOrigen;
			idMotivo = IdMotivo;
			dTotalOrigen = DTotalOrigen;
			
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int iInactivarControles = 0;
		public frmVenta(int IdCompraV, int IInactivarControles)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idCompraV = IdCompraV;
			iInactivarControles = IInactivarControles;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		string CedulaP;
		string NFactuiraP;
		bool BPagados;
		string SCedulaVendedor;
		decimal DMonto;
		int IdPagados;
		public frmVenta(string sRuc, string sNumeroFactura, bool bPagados,string sCedulaVendedor,decimal dMonto,int idPagados )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			CedulaP = sRuc;
			NFactuiraP = sNumeroFactura;
			BPagados = bPagados;
			SCedulaVendedor= sCedulaVendedor;
			DMonto = dMonto;
			IdPagados = idPagados;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public bool bRebates = false;
		private string  cedula;	
		private decimal valorR;
		private int idClienteR;
		private string clienteN;
		int idLiquidacionRebates;
		public frmVenta(string Cedula, bool BRebates, decimal ValorR,int IdCLienteR, string ClienteN,int IdLiquidacionRebates)
		{
			InitializeComponent();
			cedula = Cedula;	
			bRebates = BRebates;
			valorR = ValorR;
			idClienteR = IdCLienteR;
			clienteN = ClienteN;
			idLiquidacionRebates = IdLiquidacionRebates;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Consignacion");
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Consignacion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 8, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmVenta));
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDreccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDreccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubProyecto");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton7 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNumcuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuotas");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNumcuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
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
			this.btnEntrega = new System.Windows.Forms.Button();
			this.chkEntregaLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnImprimeED = new System.Windows.Forms.Button();
			this.dtphoracurier = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.txtresponsablecurier = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFechaEntregaCuriermbo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumTransaccion = new System.Windows.Forms.TextBox();
			this.Pieza = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lblresponsablecourier = new System.Windows.Forms.Label();
			this.optTipoVenta = new DevExpress.XtraEditors.RadioGroup();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtValorEntrega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtNumGuia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCourier = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.dtpHoraEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label67 = new System.Windows.Forms.Label();
			this.cmbDirecciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label66 = new System.Windows.Forms.Label();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtDireccionEntrega = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
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
			this.txtNumidOrigen1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidFPNC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnAnular = new System.Windows.Forms.Button();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnCliente = new System.Windows.Forms.Button();
			this.txtProyecto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumIdproyectoZoom = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnTemporal = new System.Windows.Forms.Button();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.txtIdEF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnFE = new System.Windows.Forms.Button();
			this.txtidRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
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
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirFactura = new System.Windows.Forms.MenuItem();
			this.mnuImprimirRIDE = new System.Windows.Forms.MenuItem();
			this.mnuDocsCrediPoint = new System.Windows.Forms.MenuItem();
			this.mnuImprimirGarantiaExtendida = new System.Windows.Forms.MenuItem();
			this.mnuImprimirGEP = new System.Windows.Forms.MenuItem();
			this.mnuImprimirCP = new System.Windows.Forms.MenuItem();
			this.mnuImprimirPR = new System.Windows.Forms.MenuItem();
			this.mnuImprimirASD = new System.Windows.Forms.MenuItem();
			this.mnuImprimirDYD = new System.Windows.Forms.MenuItem();
			this.mnuDocsCrediPointBlanco = new System.Windows.Forms.MenuItem();
			this.mnuImprimirContrato = new System.Windows.Forms.MenuItem();
			this.mnuImprimirNuovo = new System.Windows.Forms.MenuItem();
			this.mnuBI = new System.Windows.Forms.ContextMenu();
			this.mnuBIPedidos = new System.Windows.Forms.MenuItem();
			this.mnuBICotizaciones = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.mnuBILaboratorio = new System.Windows.Forms.MenuItem();
			this.label14 = new System.Windows.Forms.Label();
			this.txtClienteNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbBodega1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnCliente2 = new System.Windows.Forms.Button();
			this.txtRuc1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumIdCliente1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.btnValidarMail = new System.Windows.Forms.Button();
			this.txtProteccionDD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.chkFreelance = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnSubir = new System.Windows.Forms.Button();
			this.lblRecurrente = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNumContadoCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFechaPrimerPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblYupana = new System.Windows.Forms.Label();
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
			((System.ComponentModel.ISupportInitialize)(this.dtphoracurier)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtresponsablecurier)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaCuriermbo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoVenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumGuia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCourier)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionEntrega)).BeginInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidFPNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).BeginInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtClienteNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaPrimerPago)).BeginInit();
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
			this.ultraTabPageControl9.Size = new System.Drawing.Size(734, 186);
			// 
			// label72
			// 
			this.label72.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 2);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(89, 16);
			this.label72.TabIndex = 153;
			this.label72.Text = "Ingrese la Serial:";
			this.label72.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraButton1.ForeColor = System.Drawing.Color.Black;
			this.ultraButton1.Location = new System.Drawing.Point(760, 55);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 24);
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
			ultraGridColumn4.Width = 253;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 446;
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
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
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
																										 ultraGridColumn10});
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
			appearance4.FontData.SizeInPoints = 9F;
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
			this.grdSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerial.Location = new System.Drawing.Point(7, 26);
			this.grdSerial.Name = "grdSerial";
			this.grdSerial.Size = new System.Drawing.Size(720, 155);
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
			ultraDataColumn10.DataType = typeof(bool);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
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
			this.ultraTabPageControl10.Size = new System.Drawing.Size(734, 186);
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
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 39;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 22;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 22;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 44;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance9;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "%";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 38;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 52;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 138;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance10;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Width = 68;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance11;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Width = 75;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
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
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(328, 0);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(392, 182);
			this.ultraGrid2.TabIndex = 259;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
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
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
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
			this.ultraTabPageControl2.Size = new System.Drawing.Size(734, 186);
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
			this.pgEntrega.Controls.Add(this.btnEntrega);
			this.pgEntrega.Controls.Add(this.chkEntregaLocal);
			this.pgEntrega.Controls.Add(this.btnImprimeED);
			this.pgEntrega.Controls.Add(this.dtphoracurier);
			this.pgEntrega.Controls.Add(this.label9);
			this.pgEntrega.Controls.Add(this.txtresponsablecurier);
			this.pgEntrega.Controls.Add(this.dtFechaEntregaCuriermbo1);
			this.pgEntrega.Controls.Add(this.txtNumTransaccion);
			this.pgEntrega.Controls.Add(this.Pieza);
			this.pgEntrega.Controls.Add(this.label4);
			this.pgEntrega.Controls.Add(this.lblresponsablecourier);
			this.pgEntrega.Controls.Add(this.optTipoVenta);
			this.pgEntrega.Controls.Add(this.label13);
			this.pgEntrega.Controls.Add(this.label12);
			this.pgEntrega.Controls.Add(this.txtValorEntrega);
			this.pgEntrega.Controls.Add(this.label11);
			this.pgEntrega.Controls.Add(this.txtNumGuia);
			this.pgEntrega.Controls.Add(this.cmbCourier);
			this.pgEntrega.Controls.Add(this.label10);
			this.pgEntrega.Controls.Add(this.dtpHoraEntrega);
			this.pgEntrega.Controls.Add(this.label67);
			this.pgEntrega.Controls.Add(this.cmbDirecciones);
			this.pgEntrega.Controls.Add(this.label66);
			this.pgEntrega.Controls.Add(this.chkEntregado);
			this.pgEntrega.Controls.Add(this.dtFechaEntrega);
			this.pgEntrega.Controls.Add(this.label19);
			this.pgEntrega.Controls.Add(this.txtDireccionEntrega);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(734, 186);
			this.pgEntrega.Paint += new System.Windows.Forms.PaintEventHandler(this.pgEntrega_Paint);
			// 
			// btnEntrega
			// 
			this.btnEntrega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEntrega.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEntrega.CausesValidation = false;
			this.btnEntrega.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEntrega.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEntrega.Image = ((System.Drawing.Image)(resources.GetObject("btnEntrega.Image")));
			this.btnEntrega.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEntrega.Location = new System.Drawing.Point(288, 4);
			this.btnEntrega.Name = "btnEntrega";
			this.btnEntrega.Size = new System.Drawing.Size(77, 23);
			this.btnEntrega.TabIndex = 428;
			this.btnEntrega.Text = "Ubicación";
			this.btnEntrega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEntrega.Visible = false;
			this.btnEntrega.Click += new System.EventHandler(this.btnEntrega_Click);
			// 
			// chkEntregaLocal
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregaLocal.Appearance = appearance26;
			this.chkEntregaLocal.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregaLocal.Enabled = false;
			this.chkEntregaLocal.Location = new System.Drawing.Point(160, 32);
			this.chkEntregaLocal.Name = "chkEntregaLocal";
			this.chkEntregaLocal.Size = new System.Drawing.Size(120, 16);
			this.chkEntregaLocal.TabIndex = 427;
			this.chkEntregaLocal.Text = "Entrega en Local";
			this.chkEntregaLocal.CheckedChanged += new System.EventHandler(this.chkEntregaLocal_CheckedChanged);
			// 
			// btnImprimeED
			// 
			this.btnImprimeED.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimeED.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnImprimeED.CausesValidation = false;
			this.btnImprimeED.Enabled = false;
			this.btnImprimeED.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimeED.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimeED.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimeED.Image")));
			this.btnImprimeED.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimeED.Location = new System.Drawing.Point(320, 52);
			this.btnImprimeED.Name = "btnImprimeED";
			this.btnImprimeED.Size = new System.Drawing.Size(77, 23);
			this.btnImprimeED.TabIndex = 426;
			this.btnImprimeED.Text = "Imprimir";
			this.btnImprimeED.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimeED.Click += new System.EventHandler(this.btnImprimeED_Click);
			// 
			// dtphoracurier
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtphoracurier.Appearance = appearance27;
			this.dtphoracurier.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtphoracurier.DateButtons.Add(dateButton3);
			this.dtphoracurier.Enabled = false;
			this.dtphoracurier.Format = "HH:mm";
			this.dtphoracurier.Location = new System.Drawing.Point(656, 144);
			this.dtphoracurier.Name = "dtphoracurier";
			this.dtphoracurier.NonAutoSizeHeight = 23;
			this.dtphoracurier.Size = new System.Drawing.Size(80, 21);
			this.dtphoracurier.SpinButtonsVisible = true;
			this.dtphoracurier.TabIndex = 425;
			this.dtphoracurier.Value = ((object)(resources.GetObject("dtphoracurier.Value")));
			this.dtphoracurier.Visible = false;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Enabled = false;
			this.label9.Location = new System.Drawing.Point(624, 144);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(28, 16);
			this.label9.TabIndex = 424;
			this.label9.Text = "Hora";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Visible = false;
			// 
			// txtresponsablecurier
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtresponsablecurier.Appearance = appearance28;
			this.txtresponsablecurier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtresponsablecurier.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtresponsablecurier.Enabled = false;
			this.txtresponsablecurier.Location = new System.Drawing.Point(512, 72);
			this.txtresponsablecurier.MaxLength = 20;
			this.txtresponsablecurier.Name = "txtresponsablecurier";
			this.txtresponsablecurier.Size = new System.Drawing.Size(216, 22);
			this.txtresponsablecurier.TabIndex = 423;
			// 
			// dtFechaEntregaCuriermbo1
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntregaCuriermbo1.Appearance = appearance29;
			this.dtFechaEntregaCuriermbo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaEntregaCuriermbo1.DateButtons.Add(dateButton4);
			this.dtFechaEntregaCuriermbo1.Enabled = false;
			this.dtFechaEntregaCuriermbo1.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaEntregaCuriermbo1.Location = new System.Drawing.Point(512, 144);
			this.dtFechaEntregaCuriermbo1.Name = "dtFechaEntregaCuriermbo1";
			this.dtFechaEntregaCuriermbo1.NonAutoSizeHeight = 23;
			this.dtFechaEntregaCuriermbo1.Size = new System.Drawing.Size(160, 21);
			this.dtFechaEntregaCuriermbo1.SpinButtonsVisible = true;
			this.dtFechaEntregaCuriermbo1.TabIndex = 422;
			this.dtFechaEntregaCuriermbo1.Value = new System.DateTime(2018, 2, 9, 0, 0, 0, 0);
			// 
			// txtNumTransaccion
			// 
			this.txtNumTransaccion.Location = new System.Drawing.Point(688, 104);
			this.txtNumTransaccion.Name = "txtNumTransaccion";
			this.txtNumTransaccion.Size = new System.Drawing.Size(40, 20);
			this.txtNumTransaccion.TabIndex = 421;
			this.txtNumTransaccion.Text = "";
			// 
			// Pieza
			// 
			this.Pieza.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Pieza.Location = new System.Drawing.Point(640, 104);
			this.Pieza.Name = "Pieza";
			this.Pieza.Size = new System.Drawing.Size(40, 16);
			this.Pieza.TabIndex = 420;
			this.Pieza.Text = "Pieza";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(408, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(77, 16);
			this.label4.TabIndex = 418;
			this.label4.Text = "Fecha Courier:";
			// 
			// lblresponsablecourier
			// 
			this.lblresponsablecourier.AutoSize = true;
			this.lblresponsablecourier.Location = new System.Drawing.Point(408, 72);
			this.lblresponsablecourier.Name = "lblresponsablecourier";
			this.lblresponsablecourier.Size = new System.Drawing.Size(106, 16);
			this.lblresponsablecourier.TabIndex = 416;
			this.lblresponsablecourier.Text = "Responsable Courier";
			// 
			// optTipoVenta
			// 
			this.optTipoVenta.Enabled = false;
			this.optTipoVenta.Location = new System.Drawing.Point(8, 0);
			this.optTipoVenta.Name = "optTipoVenta";
			// 
			// optTipoVenta.Properties
			// 
			this.optTipoVenta.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.optTipoVenta.Properties.Appearance.Options.UseBackColor = true;
			this.optTipoVenta.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Venta Directa"),
																																																				 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Auto Consumo")});
			this.optTipoVenta.Size = new System.Drawing.Size(136, 32);
			this.optTipoVenta.TabIndex = 415;
			this.optTipoVenta.SelectedIndexChanged += new System.EventHandler(this.optTipoVenta_SelectedIndexChanged);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Enabled = false;
			this.label13.Location = new System.Drawing.Point(0, 128);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(75, 16);
			this.label13.TabIndex = 414;
			this.label13.Text = "Dirección IWY";
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(405, 100);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(89, 16);
			this.label12.TabIndex = 412;
			this.label12.Text = "Valor del Courier";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValorEntrega
			// 
			this.txtValorEntrega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorEntrega.Appearance = appearance30;
			this.txtValorEntrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorEntrega.Enabled = false;
			this.txtValorEntrega.FormatString = "#,##0.00";
			this.txtValorEntrega.Location = new System.Drawing.Point(509, 100);
			this.txtValorEntrega.Name = "txtValorEntrega";
			this.txtValorEntrega.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorEntrega.PromptChar = ' ';
			this.txtValorEntrega.Size = new System.Drawing.Size(120, 22);
			this.txtValorEntrega.TabIndex = 413;
			this.txtValorEntrega.Click += new System.EventHandler(this.txtValorEntrega_Click);
			this.txtValorEntrega.Enter += new System.EventHandler(this.txtValorEntrega_Enter);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(408, 40);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(84, 16);
			this.label11.TabIndex = 411;
			this.label11.Text = "Numero de guía";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumGuia
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumGuia.Appearance = appearance31;
			this.txtNumGuia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumGuia.Enabled = false;
			this.txtNumGuia.Location = new System.Drawing.Point(512, 40);
			this.txtNumGuia.MaxLength = 20;
			this.txtNumGuia.Name = "txtNumGuia";
			this.txtNumGuia.Size = new System.Drawing.Size(216, 22);
			this.txtNumGuia.TabIndex = 410;
			this.txtNumGuia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumGuia_KeyPress);
			// 
			// cmbCourier
			// 
			appearance32.FontData.Name = "Tahoma";
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCourier.Appearance = appearance32;
			this.cmbCourier.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCourier.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCourier.DataSource = this.ultraDataSource5;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Width = 240;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbCourier.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCourier.DisplayMember = "Nombre";
			this.cmbCourier.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCourier.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCourier.Enabled = false;
			this.cmbCourier.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCourier.Location = new System.Drawing.Point(512, 8);
			this.cmbCourier.Name = "cmbCourier";
			this.cmbCourier.Size = new System.Drawing.Size(216, 22);
			this.cmbCourier.TabIndex = 408;
			this.cmbCourier.ValueMember = "idCliente";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Enabled = false;
			this.label10.Location = new System.Drawing.Point(408, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(40, 16);
			this.label10.TabIndex = 409;
			this.label10.Text = "Courier";
			// 
			// dtpHoraEntrega
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHoraEntrega.Appearance = appearance33;
			this.dtpHoraEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtpHoraEntrega.DateButtons.Add(dateButton5);
			this.dtpHoraEntrega.Enabled = false;
			this.dtpHoraEntrega.Format = "HH:mm";
			this.dtpHoraEntrega.Location = new System.Drawing.Point(208, 56);
			this.dtpHoraEntrega.Name = "dtpHoraEntrega";
			this.dtpHoraEntrega.NonAutoSizeHeight = 23;
			this.dtpHoraEntrega.Size = new System.Drawing.Size(88, 21);
			this.dtpHoraEntrega.SpinButtonsVisible = true;
			this.dtpHoraEntrega.TabIndex = 404;
			this.dtpHoraEntrega.Value = ((object)(resources.GetObject("dtpHoraEntrega.Value")));
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Enabled = false;
			this.label67.Location = new System.Drawing.Point(176, 56);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 16);
			this.label67.TabIndex = 128;
			this.label67.Text = "Hora";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDirecciones
			// 
			appearance34.FontData.Name = "Tahoma";
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDirecciones.Appearance = appearance34;
			this.cmbDirecciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDirecciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDirecciones.DataSource = this.ultraDataSource1;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 360;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbDirecciones.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbDirecciones.DisplayMember = "Direccion";
			this.cmbDirecciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDirecciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDirecciones.Enabled = false;
			this.cmbDirecciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDirecciones.Location = new System.Drawing.Point(80, 88);
			this.cmbDirecciones.Name = "cmbDirecciones";
			this.cmbDirecciones.Size = new System.Drawing.Size(320, 22);
			this.cmbDirecciones.TabIndex = 29;
			this.cmbDirecciones.ValueMember = "idDreccion";
			this.cmbDirecciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbDirecciones_KeyPress);
			this.cmbDirecciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDirecciones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Enabled = false;
			this.label66.Location = new System.Drawing.Point(0, 88);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(50, 16);
			this.label66.TabIndex = 126;
			this.label66.Text = "Dirección";
			// 
			// chkEntregado
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance35;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.Enabled = false;
			this.chkEntregado.Location = new System.Drawing.Point(160, 8);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(120, 16);
			this.chkEntregado.TabIndex = 26;
			this.chkEntregado.Text = "Entrega a domicilio";
			this.chkEntregado.CheckedChanged += new System.EventHandler(this.chkEntregado_CheckedChanged);
			// 
			// dtFechaEntrega
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance36;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton6);
			this.dtFechaEntrega.Enabled = false;
			this.dtFechaEntrega.Format = "dd/MM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(48, 56);
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
			this.label19.Location = new System.Drawing.Point(8, 56);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(34, 16);
			this.label19.TabIndex = 44;
			this.label19.Text = "Fecha";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDireccionEntrega
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccionEntrega.Appearance = appearance37;
			this.txtDireccionEntrega.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccionEntrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccionEntrega.Enabled = false;
			this.txtDireccionEntrega.Location = new System.Drawing.Point(80, 120);
			this.txtDireccionEntrega.Multiline = true;
			this.txtDireccionEntrega.Name = "txtDireccionEntrega";
			this.txtDireccionEntrega.Size = new System.Drawing.Size(320, 56);
			this.txtDireccionEntrega.TabIndex = 277;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(734, 186);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.AcceptsTab = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance38;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(7, 9);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(717, 135);
			this.txtNotas.TabIndex = 30;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(216, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 128;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Click += new System.EventHandler(this.lblBodega_Click);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(704, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 151;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance39.FontData.Name = "Tahoma";
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance39;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 335;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(768, 7);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(250, 22);
			this.cmbVendedor.TabIndex = 3;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 35);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 16);
			this.lblIdentificacion.TabIndex = 161;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(256, 35);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 16);
			this.lblCliente.TabIndex = 160;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipoRuc
			// 
			appearance40.FontData.Name = "Tahoma";
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance40;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 233;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 32);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 5;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoRuc_KeyPress);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance41;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(352, 32);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 6;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.TextChanged += new System.EventHandler(this.txtRuc_TextChanged);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance42;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(560, 32);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(256, 22);
			this.txtNombre.TabIndex = 7;
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(496, 83);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(43, 16);
			this.label61.TabIndex = 164;
			this.label61.Text = "Entrada";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label61.Click += new System.EventHandler(this.label61_Click);
			// 
			// txtCuotaDeEntrada
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDeEntrada.Appearance = appearance43;
			this.txtCuotaDeEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDeEntrada.Enabled = false;
			this.txtCuotaDeEntrada.FormatString = "#,##0.00";
			this.txtCuotaDeEntrada.Location = new System.Drawing.Point(560, 80);
			this.txtCuotaDeEntrada.Name = "txtCuotaDeEntrada";
			this.txtCuotaDeEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDeEntrada.PromptChar = ' ';
			this.txtCuotaDeEntrada.Size = new System.Drawing.Size(80, 22);
			this.txtCuotaDeEntrada.TabIndex = 11;
			this.txtCuotaDeEntrada.ValidationError += new Infragistics.Win.UltraWinEditors.UltraNumericEditorBase.ValidationErrorEventHandler(this.txtCuotaDeEntrada_ValidationError);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 167;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance44.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance44;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 56;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 66;
			ultraGridColumn34.Header.VisiblePosition = 15;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn35.Header.VisiblePosition = 3;
			ultraGridColumn35.Width = 249;
			ultraGridColumn36.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 4;
			ultraGridColumn36.Width = 313;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance45;
			ultraGridColumn37.Format = "";
			ultraGridColumn37.Header.VisiblePosition = 5;
			ultraGridColumn37.PromptChar = ' ';
			ultraGridColumn37.Width = 79;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance46;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 6;
			ultraGridColumn38.PromptChar = ' ';
			ultraGridColumn38.Width = 97;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance47;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "IVA";
			ultraGridColumn39.Header.VisiblePosition = 7;
			ultraGridColumn39.PromptChar = ' ';
			ultraGridColumn39.Width = 60;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance48;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.PromptChar = ' ';
			ultraGridColumn40.Width = 90;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance49;
			ultraGridColumn41.Format = "#,##0.0000";
			ultraGridColumn41.Header.Caption = "% Desc.";
			ultraGridColumn41.Header.VisiblePosition = 9;
			ultraGridColumn41.PromptChar = ' ';
			ultraGridColumn41.Width = 68;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance50;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Subtotal";
			ultraGridColumn42.Header.VisiblePosition = 10;
			ultraGridColumn42.PromptChar = ' ';
			ultraGridColumn42.Width = 120;
			ultraGridColumn43.Header.VisiblePosition = 14;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 12;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 11;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 16;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 63;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 17;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 18;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 19;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.VisiblePosition = 20;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 21;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 22;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 94;
			ultraGridColumn53.Header.VisiblePosition = 23;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 93;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53});
			ultraGridBand7.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance51.ForeColor = System.Drawing.Color.Black;
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance51;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance52.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance52;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance53.FontData.BoldAsString = "True";
			appearance53.FontData.Name = "Arial";
			appearance53.FontData.SizeInPoints = 9F;
			appearance53.ForeColor = System.Drawing.Color.White;
			appearance53.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance53;
			appearance54.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance54.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance56;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(9, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1007, 158);
			this.ultraGrid1.TabIndex = 14;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn25.DefaultValue = 0;
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn26.DefaultValue = 0;
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn27.DefaultValue = 0;
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn28.DefaultValue = 0;
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn29.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn32.DefaultValue = 0;
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn33.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn34.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn35.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn36.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(bool);
			ultraDataColumn38.DefaultValue = false;
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn39.DefaultValue = 0;
			ultraDataColumn40.DataType = typeof(bool);
			ultraDataColumn40.DefaultValue = false;
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn41.DefaultValue = 0;
			ultraDataColumn42.DataType = typeof(bool);
			ultraDataColumn42.DefaultValue = false;
			ultraDataColumn43.DataType = typeof(bool);
			ultraDataColumn43.DefaultValue = false;
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn44.DefaultValue = 0;
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn45.DefaultValue = 0;
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48});
			this.ultraDataSource2.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource2_CellDataRequested);
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(8, 83);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 190;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance57;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(96, 80);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 4;
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
			this.lblSubTotal.Location = new System.Drawing.Point(752, 340);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(46, 16);
			this.lblSubTotal.TabIndex = 208;
			this.lblSubTotal.Text = "Subtotal";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(752, 411);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 206;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(752, 363);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(58, 16);
			this.lblDescuento.TabIndex = 204;
			this.lblDescuento.Text = "Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva0.Location = new System.Drawing.Point(888, 312);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 200;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(992, 312);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 16);
			this.label55.TabIndex = 203;
			this.label55.Text = "IVA";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance58;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(840, 432);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 34);
			this.txtTotal.TabIndex = 222;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance59;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(936, 409);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 221;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance60;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(840, 361);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 219;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance61;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(936, 361);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 216;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance62;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(936, 337);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 214;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance63;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(840, 337);
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
			this.txtNumIdAsiento.Location = new System.Drawing.Point(952, 488);
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
			this.label1.Location = new System.Drawing.Point(752, 442);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 223;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance64;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.pgEntrega);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl9);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl10);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl2);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 281);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(736, 207);
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
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(734, 186);
			// 
			// cmbFormaPago
			// 
			appearance65.FontData.Name = "Tahoma";
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance65;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 1;
			ultraGridColumn55.Width = 152;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 56);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 22);
			this.cmbFormaPago.TabIndex = 10;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 59);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 225;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtComprobante
			// 
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance66;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(352, 80);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(104, 22);
			this.txtComprobante.TabIndex = 9;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(256, 83);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 226;
			this.lblComprobante.Text = "Comprobante";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.txtNumIdCliente.Location = new System.Drawing.Point(936, 488);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 228;
			this.txtNumIdCliente.Visible = false;
			this.txtNumIdCliente.ValueChanged += new System.EventHandler(this.txtNumIdCliente_ValueChanged);
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(904, 488);
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
			this.txtNumidOrigen.Location = new System.Drawing.Point(920, 488);
			this.txtNumidOrigen.Name = "txtNumidOrigen";
			this.txtNumidOrigen.PromptChar = ' ';
			this.txtNumidOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen.TabIndex = 230;
			this.txtNumidOrigen.Visible = false;
			// 
			// txtNumidOrigen1
			// 
			this.txtNumidOrigen1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen1.Enabled = false;
			this.txtNumidOrigen1.Location = new System.Drawing.Point(984, 488);
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
			this.txtNumidFPNC.Location = new System.Drawing.Point(1000, 488);
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
			this.txtIdSucursal.Location = new System.Drawing.Point(904, 513);
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
			this.txtBodegaPredef.Location = new System.Drawing.Point(920, 513);
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
			this.txtIdTipoFactura.Location = new System.Drawing.Point(936, 513);
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
			this.txtNumEstado.Location = new System.Drawing.Point(952, 513);
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
			this.txtIdComprobante.Location = new System.Drawing.Point(968, 513);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 238;
			this.txtIdComprobante.Visible = false;
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(984, 513);
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
			this.lblEstado.Location = new System.Drawing.Point(648, 78);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 240;
			// 
			// cmbBodega
			// 
			appearance67.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance67;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn56.Header.VisiblePosition = 2;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Width = 180;
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(256, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 2;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.AfterCloseUp += new System.EventHandler(this.cmbBodega_AfterCloseUp);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(496, 35);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(38, 16);
			this.label3.TabIndex = 244;
			this.label3.Text = "Cliente";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(556, 505);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 36;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(752, 312);
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
			this.btnBuscar.Location = new System.Drawing.Point(708, 505);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 38;
			this.btnBuscar.Text = "Documentos";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(824, 32);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 8;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// txtProyecto
			// 
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyecto.Appearance = appearance68;
			this.txtProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyecto.Enabled = false;
			this.txtProyecto.Location = new System.Drawing.Point(800, 472);
			this.txtProyecto.Name = "txtProyecto";
			this.txtProyecto.Size = new System.Drawing.Size(32, 22);
			this.txtProyecto.TabIndex = 12;
			this.txtProyecto.Visible = false;
			this.txtProyecto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProyecto_KeyDown);
			this.txtProyecto.Validating += new System.ComponentModel.CancelEventHandler(this.txtProyecto_Validating);
			this.txtProyecto.Validated += new System.EventHandler(this.txtProyecto_Validated);
			this.txtProyecto.ValueChanged += new System.EventHandler(this.txtProyecto_ValueChanged);
			// 
			// txtNumIdproyectoZoom
			// 
			this.txtNumIdproyectoZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdproyectoZoom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdproyectoZoom.Enabled = false;
			this.txtNumIdproyectoZoom.Location = new System.Drawing.Point(1000, 513);
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
			this.btnTemporal.Location = new System.Drawing.Point(397, 505);
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
			this.btnAsiento.Location = new System.Drawing.Point(474, 505);
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
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 488);
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
			this.txtIdEF.Location = new System.Drawing.Point(888, 488);
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
			this.btnFE.Location = new System.Drawing.Point(904, 505);
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
			this.txtidRetencion.Location = new System.Drawing.Point(888, 513);
			this.txtidRetencion.Name = "txtidRetencion";
			this.txtidRetencion.PromptChar = ' ';
			this.txtidRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtidRetencion.TabIndex = 266;
			this.txtidRetencion.Visible = false;
			// 
			// btnClaveDescuento
			// 
			this.btnClaveDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnClaveDescuento.CausesValidation = false;
			this.btnClaveDescuento.Enabled = false;
			this.btnClaveDescuento.Location = new System.Drawing.Point(867, 505);
			this.btnClaveDescuento.Name = "btnClaveDescuento";
			this.btnClaveDescuento.Size = new System.Drawing.Size(24, 23);
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
			this.txtidBanco.Location = new System.Drawing.Point(744, 465);
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
			this.txtidPlan.Location = new System.Drawing.Point(776, 465);
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
			this.txtidTarjeta.Location = new System.Drawing.Point(760, 465);
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
			this.txtidPlazo.Location = new System.Drawing.Point(792, 465);
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
			this.label5.Location = new System.Drawing.Point(752, 388);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(174, 16);
			this.label5.TabIndex = 273;
			this.label5.Text = "BASE IMPONIBLE RETENCION";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBaseImponible
			// 
			this.txtBaseImponible.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBaseImponible.Appearance = appearance69;
			this.txtBaseImponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBaseImponible.Enabled = false;
			this.txtBaseImponible.FormatString = "#,##0.00";
			this.txtBaseImponible.Location = new System.Drawing.Point(936, 384);
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
			this.lblPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPlan.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPlan.Location = new System.Drawing.Point(752, 473);
			this.lblPlan.Name = "lblPlan";
			this.lblPlan.Size = new System.Drawing.Size(0, 14);
			this.lblPlan.TabIndex = 275;
			this.lblPlan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance70;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton7.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton7);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
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
			this.txtIdAsientoRetencion.Location = new System.Drawing.Point(808, 465);
			this.txtIdAsientoRetencion.Name = "txtIdAsientoRetencion";
			this.txtIdAsientoRetencion.PromptChar = ' ';
			this.txtIdAsientoRetencion.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsientoRetencion.TabIndex = 276;
			this.txtIdAsientoRetencion.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 505);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(82, 505);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(162, 505);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 33;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(240, 505);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 32;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(319, 505);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 31;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(634, 505);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 37;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(790, 505);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 39;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirFactura,
																																								this.mnuImprimirRIDE,
																																								this.mnuDocsCrediPoint,
																																								this.mnuImprimirGarantiaExtendida,
																																								this.mnuImprimirGEP,
																																								this.mnuImprimirCP,
																																								this.mnuImprimirPR,
																																								this.mnuImprimirASD,
																																								this.mnuImprimirDYD,
																																								this.mnuDocsCrediPointBlanco,
																																								this.mnuImprimirContrato,
																																								this.mnuImprimirNuovo});
			this.mnuImprimir.Popup += new System.EventHandler(this.mnuImprimir_Popup);
			// 
			// mnuImprimirFactura
			// 
			this.mnuImprimirFactura.Enabled = false;
			this.mnuImprimirFactura.Index = 0;
			this.mnuImprimirFactura.Text = "Factura";
			this.mnuImprimirFactura.Click += new System.EventHandler(this.mnuImprimirFactura_Click);
			// 
			// mnuImprimirRIDE
			// 
			this.mnuImprimirRIDE.Enabled = false;
			this.mnuImprimirRIDE.Index = 1;
			this.mnuImprimirRIDE.Text = "RIDE";
			this.mnuImprimirRIDE.Click += new System.EventHandler(this.mnuImprimirRIDE_Click);
			// 
			// mnuDocsCrediPoint
			// 
			this.mnuDocsCrediPoint.Enabled = false;
			this.mnuDocsCrediPoint.Index = 2;
			this.mnuDocsCrediPoint.Text = "Documentos de CrediPOINT";
			this.mnuDocsCrediPoint.Click += new System.EventHandler(this.mnuDocsCrediPoint_Click);
			// 
			// mnuImprimirGarantiaExtendida
			// 
			this.mnuImprimirGarantiaExtendida.Enabled = false;
			this.mnuImprimirGarantiaExtendida.Index = 3;
			this.mnuImprimirGarantiaExtendida.Text = "Garantia Extendida";
			this.mnuImprimirGarantiaExtendida.Click += new System.EventHandler(this.mnuImprimirGarantiaExtendida_Click);
			// 
			// mnuImprimirGEP
			// 
			this.mnuImprimirGEP.Enabled = false;
			this.mnuImprimirGEP.Index = 4;
			this.mnuImprimirGEP.Text = "Garantia Extendida Plus";
			this.mnuImprimirGEP.Click += new System.EventHandler(this.mnuImprimirGEP_Click);
			// 
			// mnuImprimirCP
			// 
			this.mnuImprimirCP.Enabled = false;
			this.mnuImprimirCP.Index = 5;
			this.mnuImprimirCP.Text = "Compra Protegida";
			this.mnuImprimirCP.Click += new System.EventHandler(this.mnuImprimirCP_Click);
			// 
			// mnuImprimirPR
			// 
			this.mnuImprimirPR.Enabled = false;
			this.mnuImprimirPR.Index = 6;
			this.mnuImprimirPR.Text = "Point Respaldo";
			this.mnuImprimirPR.Click += new System.EventHandler(this.mnuImprimirPR_Click);
			// 
			// mnuImprimirASD
			// 
			this.mnuImprimirASD.Enabled = false;
			this.mnuImprimirASD.Index = 7;
			this.mnuImprimirASD.Text = "Asistencia Salud Dental";
			this.mnuImprimirASD.Click += new System.EventHandler(this.mnuImprimirASD_Click);
			// 
			// mnuImprimirDYD
			// 
			this.mnuImprimirDYD.Enabled = false;
			this.mnuImprimirDYD.Index = 8;
			this.mnuImprimirDYD.Text = "Seguro D&FI";
			this.mnuImprimirDYD.Click += new System.EventHandler(this.mnuImprimirDYD_Click);
			// 
			// mnuDocsCrediPointBlanco
			// 
			this.mnuDocsCrediPointBlanco.Enabled = false;
			this.mnuDocsCrediPointBlanco.Index = 9;
			this.mnuDocsCrediPointBlanco.Text = "Documentos de CrediPoint en Blanco";
			this.mnuDocsCrediPointBlanco.Click += new System.EventHandler(this.mnuDocsCrediPointBlanco_Click);
			// 
			// mnuImprimirContrato
			// 
			this.mnuImprimirContrato.Enabled = false;
			this.mnuImprimirContrato.Index = 10;
			this.mnuImprimirContrato.Text = "Contrato Celulares";
			this.mnuImprimirContrato.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// mnuImprimirNuovo
			// 
			this.mnuImprimirNuovo.Enabled = false;
			this.mnuImprimirNuovo.Index = 11;
			this.mnuImprimirNuovo.Text = "Contrato Nuovo";
			this.mnuImprimirNuovo.Click += new System.EventHandler(this.mnuImprimirNuovo_Click);
			// 
			// mnuBI
			// 
			this.mnuBI.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					this.mnuBIPedidos,
																																					this.mnuBICotizaciones,
																																					this.menuItem1,
																																					this.mnuBILaboratorio});
			// 
			// mnuBIPedidos
			// 
			this.mnuBIPedidos.Index = 0;
			this.mnuBIPedidos.Text = "Pedidos";
			this.mnuBIPedidos.Click += new System.EventHandler(this.mnuBIPedidos_Click);
			// 
			// mnuBICotizaciones
			// 
			this.mnuBICotizaciones.Index = 1;
			this.mnuBICotizaciones.Text = "Cotizaciones";
			this.mnuBICotizaciones.Click += new System.EventHandler(this.mnuBICotizaciones_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Proforma";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// mnuBILaboratorio
			// 
			this.mnuBILaboratorio.Index = 3;
			this.mnuBILaboratorio.Text = "Laboratorio";
			this.mnuBILaboratorio.Click += new System.EventHandler(this.mnuBILaboratorio_Click);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(664, 88);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(95, 16);
			this.label14.TabIndex = 417;
			this.label14.Text = "N. Identificación 1";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Visible = false;
			// 
			// txtClienteNombre
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtClienteNombre.Appearance = appearance71;
			this.txtClienteNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteNombre.Enabled = false;
			this.txtClienteNombre.Location = new System.Drawing.Point(792, 56);
			this.txtClienteNombre.Name = "txtClienteNombre";
			this.txtClienteNombre.Size = new System.Drawing.Size(24, 22);
			this.txtClienteNombre.TabIndex = 416;
			this.txtClienteNombre.Visible = false;
			this.txtClienteNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClienteNombre_KeyDown);
			this.txtClienteNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtClienteNombre_Validating);
			this.txtClienteNombre.Validated += new System.EventHandler(this.txtClienteNombre_Validated);
			this.txtClienteNombre.ValueChanged += new System.EventHandler(this.txtClienteNombre_ValueChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(456, 10);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(41, 16);
			this.label15.TabIndex = 424;
			this.label15.Text = "Bodega";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega1
			// 
			this.cmbBodega1.CausesValidation = false;
			this.cmbBodega1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega1.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn59.Header.VisiblePosition = 2;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Width = 180;
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn59,
																											ultraGridColumn60,
																											ultraGridColumn61});
			this.cmbBodega1.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbBodega1.DisplayMember = "Nombre";
			this.cmbBodega1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega1.Enabled = false;
			this.cmbBodega1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega1.Location = new System.Drawing.Point(512, 7);
			this.cmbBodega1.MaxDropDownItems = 30;
			this.cmbBodega1.Name = "cmbBodega1";
			this.cmbBodega1.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega1.TabIndex = 423;
			this.cmbBodega1.ValueMember = "Bodega";
			// 
			// btnCliente2
			// 
			this.btnCliente2.Enabled = false;
			this.btnCliente2.Location = new System.Drawing.Point(824, 56);
			this.btnCliente2.Name = "btnCliente2";
			this.btnCliente2.Size = new System.Drawing.Size(24, 22);
			this.btnCliente2.TabIndex = 425;
			this.btnCliente2.Text = "...";
			this.btnCliente2.Visible = false;
			this.btnCliente2.Click += new System.EventHandler(this.btnCliente2_Click);
			// 
			// txtRuc1
			// 
			this.txtRuc1.AcceptsTab = true;
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc1.Appearance = appearance72;
			this.txtRuc1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc1.Enabled = false;
			this.txtRuc1.Location = new System.Drawing.Point(776, 80);
			this.txtRuc1.Name = "txtRuc1";
			this.txtRuc1.Size = new System.Drawing.Size(16, 22);
			this.txtRuc1.TabIndex = 426;
			this.txtRuc1.Visible = false;
			this.txtRuc1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc1_KeyDown);
			this.txtRuc1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc1_KeyPress);
			this.txtRuc1.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc1_Validating);
			this.txtRuc1.Validated += new System.EventHandler(this.txtRuc1_Validated);
			this.txtRuc1.ValueChanged += new System.EventHandler(this.txtRuc1_ValueChanged);
			// 
			// txtNumIdCliente1
			// 
			this.txtNumIdCliente1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente1.Enabled = false;
			this.txtNumIdCliente1.Location = new System.Drawing.Point(920, 473);
			this.txtNumIdCliente1.Name = "txtNumIdCliente1";
			this.txtNumIdCliente1.PromptChar = ' ';
			this.txtNumIdCliente1.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente1.TabIndex = 427;
			this.txtNumIdCliente1.Visible = false;
			this.txtNumIdCliente1.ValueChanged += new System.EventHandler(this.txtNumIdCliente1_ValueChanged);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(736, 59);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(48, 16);
			this.label16.TabIndex = 428;
			this.label16.Text = "Cliente 1";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Visible = false;
			// 
			// btnValidarMail
			// 
			this.btnValidarMail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnValidarMail.Enabled = false;
			this.btnValidarMail.Location = new System.Drawing.Point(856, 32);
			this.btnValidarMail.Name = "btnValidarMail";
			this.btnValidarMail.Size = new System.Drawing.Size(136, 22);
			this.btnValidarMail.TabIndex = 446;
			this.btnValidarMail.Text = "Validar Mail Promocional";
			this.btnValidarMail.Visible = false;
			this.btnValidarMail.Click += new System.EventHandler(this.btnValidarMail_Click);
			// 
			// txtProteccionDD
			// 
			appearance73.FontData.SizeInPoints = 14F;
			appearance73.ForeColor = System.Drawing.Color.Firebrick;
			appearance73.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProteccionDD.Appearance = appearance73;
			this.txtProteccionDD.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProteccionDD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProteccionDD.Enabled = false;
			this.txtProteccionDD.FormatString = "#,##0.00";
			this.txtProteccionDD.Location = new System.Drawing.Point(944, 80);
			this.txtProteccionDD.Name = "txtProteccionDD";
			this.txtProteccionDD.NullText = "0.00";
			this.txtProteccionDD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProteccionDD.PromptChar = ' ';
			this.txtProteccionDD.Size = new System.Drawing.Size(8, 30);
			this.txtProteccionDD.TabIndex = 449;
			this.txtProteccionDD.Visible = false;
			// 
			// timer1
			// 
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// chkFreelance
			// 
			this.chkFreelance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance74.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFreelance.Appearance = appearance74;
			this.chkFreelance.BackColor = System.Drawing.Color.Transparent;
			this.chkFreelance.Enabled = false;
			this.chkFreelance.Location = new System.Drawing.Point(856, 32);
			this.chkFreelance.Name = "chkFreelance";
			this.chkFreelance.Size = new System.Drawing.Size(88, 22);
			this.chkFreelance.TabIndex = 450;
			this.chkFreelance.Text = "Freelance";
			this.chkFreelance.Visible = false;
			// 
			// btnSubir
			// 
			this.btnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSubir.CausesValidation = false;
			this.btnSubir.Image = ((System.Drawing.Image)(resources.GetObject("btnSubir.Image")));
			this.btnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubir.Location = new System.Drawing.Point(893, 505);
			this.btnSubir.Name = "btnSubir";
			this.btnSubir.Size = new System.Drawing.Size(73, 23);
			this.btnSubir.TabIndex = 451;
			this.btnSubir.Text = "&Subir";
			this.btnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubir.Click += new System.EventHandler(this.btnSubir_Click);
			// 
			// lblRecurrente
			// 
			this.lblRecurrente.AutoSize = true;
			this.lblRecurrente.BackColor = System.Drawing.Color.Transparent;
			this.lblRecurrente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblRecurrente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblRecurrente.ForeColor = System.Drawing.Color.Firebrick;
			this.lblRecurrente.Location = new System.Drawing.Point(853, 56);
			this.lblRecurrente.Name = "lblRecurrente";
			this.lblRecurrente.Size = new System.Drawing.Size(0, 20);
			this.lblRecurrente.TabIndex = 713;
			this.lblRecurrente.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(256, 56);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(38, 16);
			this.label17.TabIndex = 714;
			this.label17.Text = "Cuotas";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance75.ForeColor = System.Drawing.Color.Black;
			this.txtNumContadoCredito.Appearance = appearance75;
			this.txtNumContadoCredito.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtNumContadoCredito.ForeColor = System.Drawing.Color.Black;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(936, 472);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 231;
			this.txtNumContadoCredito.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn49,
																																 ultraDataColumn50});
			// 
			// cmbNumContadoCredito
			// 
			appearance76.FontData.Name = "Tahoma";
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNumContadoCredito.Appearance = appearance76;
			this.cmbNumContadoCredito.CausesValidation = false;
			this.cmbNumContadoCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNumContadoCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNumContadoCredito.DataSource = this.ultraDataSource6;
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.Caption = "Cuotas";
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn62,
																											ultraGridColumn63});
			this.cmbNumContadoCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbNumContadoCredito.DisplayMember = "NumeroCuotas";
			this.cmbNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNumContadoCredito.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNumContadoCredito.Enabled = false;
			this.cmbNumContadoCredito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNumContadoCredito.Location = new System.Drawing.Point(352, 56);
			this.cmbNumContadoCredito.Name = "cmbNumContadoCredito";
			this.cmbNumContadoCredito.Size = new System.Drawing.Size(128, 22);
			this.cmbNumContadoCredito.TabIndex = 717;
			this.cmbNumContadoCredito.ValueMember = "idNumcuotas";
			this.cmbNumContadoCredito.ValueChanged += new System.EventHandler(this.cmbNumContadoCredito_ValueChanged);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(496, 56);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(64, 16);
			this.label18.TabIndex = 719;
			this.label18.Text = "Día de pago";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn51.DataType = typeof(System.DateTime);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn51});
			// 
			// cmbFechaPrimerPago
			// 
			this.cmbFechaPrimerPago.CausesValidation = false;
			this.cmbFechaPrimerPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFechaPrimerPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFechaPrimerPago.DataSource = this.ultraDataSource7;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Width = 100;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn64});
			this.cmbFechaPrimerPago.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbFechaPrimerPago.DisplayMember = "Fecha";
			this.cmbFechaPrimerPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFechaPrimerPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFechaPrimerPago.Enabled = false;
			this.cmbFechaPrimerPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFechaPrimerPago.Location = new System.Drawing.Point(560, 56);
			this.cmbFechaPrimerPago.MaxDropDownItems = 30;
			this.cmbFechaPrimerPago.Name = "cmbFechaPrimerPago";
			this.cmbFechaPrimerPago.Size = new System.Drawing.Size(168, 22);
			this.cmbFechaPrimerPago.TabIndex = 721;
			this.cmbFechaPrimerPago.ValueMember = "Fecha";
			this.cmbFechaPrimerPago.ValueChanged += new System.EventHandler(this.cmbFechaPrimerPago_ValueChanged);
			// 
			// lblYupana
			// 
			this.lblYupana.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblYupana.AutoSize = true;
			this.lblYupana.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblYupana.ForeColor = System.Drawing.Color.Firebrick;
			this.lblYupana.Location = new System.Drawing.Point(760, 280);
			this.lblYupana.Name = "lblYupana";
			this.lblYupana.Size = new System.Drawing.Size(0, 17);
			this.lblYupana.TabIndex = 722;
			this.lblYupana.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmVenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1022, 532);
			this.Controls.Add(this.lblYupana);
			this.Controls.Add(this.cmbFechaPrimerPago);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbNumContadoCredito);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.lblRecurrente);
			this.Controls.Add(this.btnSubir);
			this.Controls.Add(this.chkFreelance);
			this.Controls.Add(this.txtProteccionDD);
			this.Controls.Add(this.btnValidarMail);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label5);
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
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNumIdCliente1);
			this.Controls.Add(this.txtRuc1);
			this.Controls.Add(this.btnCliente2);
			this.Controls.Add(this.cmbBodega1);
			this.Controls.Add(this.txtIdAsientoRetencion);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblPlan);
			this.Controls.Add(this.txtBaseImponible);
			this.Controls.Add(this.txtidBanco);
			this.Controls.Add(this.txtidPlan);
			this.Controls.Add(this.txtidTarjeta);
			this.Controls.Add(this.txtidPlazo);
			this.Controls.Add(this.btnClaveDescuento);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.txtidRetencion);
			this.Controls.Add(this.btnFE);
			this.Controls.Add(this.txtIdEF);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnTemporal);
			this.Controls.Add(this.txtNumIdproyectoZoom);
			this.Controls.Add(this.txtProyecto);
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
			this.Controls.Add(this.txtClienteNombre);
			this.KeyPreview = true;
			this.Name = "frmVenta";
			this.Text = "VENTAS";
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
			((System.ComponentModel.ISupportInitialize)(this.dtphoracurier)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtresponsablecurier)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntregaCuriermbo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoVenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumGuia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCourier)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionEntrega)).EndInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidFPNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).EndInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtClienteNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaPrimerPago)).EndInit();
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

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}
		
		bool bPromociones = false;
		bool bExigeVendedor = false;
		bool bEmpleado = false;

		private void NuevoNotasDeCredito(int idCompraO)
		{	
			#region Maestro
			int iBodega = 0;
			int iBodegaInv = 0;
				
			string sSQL = string.Format("Exec VentaConsultaIndividual {0}", idCompraO);
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
				this.cmbTipoRuc.Value = dr.GetInt32(15);
				this.txtNumIdCliente.Value = dr.GetInt32(16);
				this.txtRuc.Text = dr.GetValue(17).ToString();
				this.txtNombre.Text = dr.GetValue(18).ToString();
				this.txtSerie.Text = dr.GetValue(20).ToString();				
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
				bVeriricador = dr.GetBoolean(56);
				this.optTipoVenta.EditValue = dr.GetInt32(57);
				this.txtNumTransaccion.Text = dr.GetString(58);
				this.txtDireccionEntrega.Text = dr.GetString(59);
				this.cmbCourier.Value = dr.GetInt32(60); 
				this.txtNumGuia.Text = dr.GetString(61);
				this.txtValorEntrega.Value = dr.GetDecimal(62);
				this.txtClienteNombre.Text = dr.GetString(63);
				iBodegaInv = dr.GetInt32(64);				
			}
			dr.Close();

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			bNuevo = true;

			this.cmbBodega.Value = iBodega;
			this.cmbBodega1.Value = iBodegaInv;

			this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtNumIdCliente.Value, cdsSeteoF);		
			#endregion Maestro
			
			#region S.R.I. si es registro nuevo
			//			SqlDataReader drSRI = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico, Promociones, ValidaTipoVenta From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
			//			drSRI.Read();
			//			if (drSRI.HasRows)
			//			{
			//				this.txtSerie.Text = drSRI.GetValue(0).ToString();				
			//				if (drSRI.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = drSRI.GetDateTime(2);
			//				bNumeracionAutomatica = drSRI.GetBoolean(3);
			//				bPromociones = drSRI.GetBoolean(4);
			//				bValidaTipoVenta = drSRI.GetBoolean(5);
			//			}
			//			drSRI.Close();	
			//					
			//			if (!bNumeracionAutomatica && bEdicion)	
			//				this.txtNumero.Enabled = true; 
			//			else
			//			{
			//				this.txtNumero.Text = "";
			//				this.txtNumero.Enabled = false;
			//			}		
			#endregion S.R.I. si es registro nuevo				

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", idCompraO);
			idBloqueaTransacciones1 = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			BUsoNotaDeCredito = true;
			
			bEdicion = true;
			
			this.txtNumContadoCredito.Value = 1;			
			this.optTipoVenta.EditValue = 1;						
									
			#region Habilita controles	
			//			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");		
			//			this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec RefrescaSerialVenta 0");	

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			FuncionesProcedimientos.EstadoGrids(true, this.grdSerial);

			this.txtSerial.Enabled = true;
			this.chkEntregado.Enabled = true;
			this.txtNotas.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			if (miAcceso.CambiarDescuento) this.btnClaveDescuento.Enabled = true;

			this.btnCliente.Enabled = true;
			this.btnBuscar.Enabled = false;		

			#endregion Habilita controles
			
			this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
			this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
		}

		private void frmVenta_Load(object sender, System.EventArgs e)
		{	
			miAcceso = new Acceso(cdsSeteoF, "0703");
			
			if (!Funcion.Permiso("118", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Ventas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			int i = MenuLatinium.iNivel;

			string carpeta_1 = "";
			carpeta_1 = @"C:\Latinium\Reportes\CompraProtegida_1.rpt";

			if (!miAcceso.BAutorizarPedido) this.mnuBILaboratorio.Visible = false;
												
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
				
			if ((bool) drSeteoF["IncluyeIva"]) bIncluyeIva = true;
			bDescuentoLineal = (bool) drSeteoF["DescuentoLineal"];
			
			bExigeVendedor = Funcion.bEscalarSQL(cdsSeteoF, "SELECT ExigeVendedor FROM SeteoF", true);

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga 1, 0");
			this.cmbNumContadoCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 101, 0");
			this.cmbFechaPrimerPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Cre_SolicitudYupanaDiaPago 35, 0");
			this.cmbCourier.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCliente, Ruc, Nombre From Cliente Where Proveedor = 1 And Courier = 1 Order By Nombre");

			HabilitaTab(false);

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if ((int)this.txtBodegaPredef.Value > 1)
			{
				this.utcTransaccion.Tabs[2].Visible = false;
				this.utcTransaccion.Tabs[3].Visible = false;
			}

			if ((int)this.txtBodegaPredef.Value == 69) this.utcTransaccion.Tabs[3].Visible = true;

			this.txtIdTipoFactura.Value = 1;
			this.txtIdComprobante.Value = 15;
				
			iDiasPosteriores = Funcion.iEscalarSQL(cdsSeteoF, "Select DiasPosteriores From SeteoF");

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(iDiasPosteriores);			

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
			if (miAcceso.CambiarDescuento) this.btnClaveDescuento.Visible = true;
			if (miAcceso.BEditaCantidad) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;

			if ((bool)drSeteoF["CotizaServicios"] && !(bool)drSeteoF["CotizaInventario"])  
				this.utcTransaccion.Tabs[0].Visible = false;								
			
			if (idCompraV > 0) this.Consultar(idCompraV);	

			if (idCompraOrigen > 0) 
			{
				this.btnCancelar_Click(sender, e);
			
				//				this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec RefrescaSerialVenta 0");	
				//				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");

				NuevoNotasDeCredito(idCompraOrigen);
			}

			if (BPagados)
				EjecutarProcesoPagados();
			if (bRebates)
			{
				EjecutarProcesoRebates();
			}

			#region Asiento Automatico Empresas
			if (bAutoConsumo)	this.btnEditar_Click(sender, e);

			if (MenuLatinium.stIdDB != 1) bAsientoAutomatico = true;
			#endregion Asiento Automatico Empresas
			
			//Para reporte venta promociones
			ReporteVentaPromocionsControles();	
		}

		private void ReporteVentaPromocionsControles()
		{
			if(iInactivarControles == 1)
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnBuscar.Enabled = false;
				this.btnTemporal.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnClaveDescuento.Enabled = false;
				this.btnGuardar.Enabled = false;
			}
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				if (!bNumeracionAutomatica) this.txtNumero.Focus(); else if (this.txtRuc.Enabled) this.txtRuc.Focus(); else
				{						
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];						
				}
		}

		private void cmbTipoRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 				
			{
				if (this.cmbTipoRuc.ActiveRow != null)
				{
					if ((int)this.cmbTipoRuc.Value == 11)
					{				
						this.txtClienteNombre.Enabled=true;
						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];	
						
					}
					else this.txtRuc.Focus();				
				}
			}
		}

		private void EstadoRetencion(bool bEstado)
		{
			this.btnRetener.Enabled = bEstado;
			this.btnCancelarRetencion.Enabled = bEstado;
			this.txtSerieRetencion.Enabled = bEstado;
			this.txtNumeroRetencion.Enabled = bEstado;
			this.txtAutorizacionRetencion.Enabled = bEstado;
			this.dtFechaRetencion.Enabled = bEstado;

			if (bEstado) 
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				this.dtFechaRetencion.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
				this.dtFechaRetencion.CalendarInfo.MaxDate = dtFecha.AddDays(5);			
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5 || (int)this.cmbFormaPago.Value != 19) Funcion.CamposNumericos(sender, e);
			if ((int)this.cmbTipoRuc.Value == 6 ) Funcion.CamposNumericosPasaporte(sender, e);

			if (e.KeyChar == 13)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];				
			}
		}
		private void txtRuc1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//if ((int)this.cmbTipoRuc.Value == 11) Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];				
			}
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{	
			if (bEdicion)
			{
				this.txtRuc1.Enabled = false;
				this.txtClienteNombre.Enabled = false;

				if (bNuevo && (int)this.cmbFormaPago.Value != 4)
				{
					this.txtRuc.Enabled = ((int)this.cmbFormaPago.Value != 35);
					this.txtNumIdCliente.Value = 0;
					this.txtRuc.Value = "";
					this.txtNombre.Value = "";
				}

				if (this.cmbTipoRuc.ActiveRow != null)
				{
					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

					#region Consumidor Final
					if ((int)this.cmbTipoRuc.Value == 11) 
					{						
						this.txtRuc.Enabled = false;
						//this.txtRuc1.Enabled = true;
						this.txtRuc.MaxLength = 13;
						this.btnCliente.Enabled = false;
						this.btnCliente2.Enabled = true;

						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select cr.idCliente, cl.Ruc, cl.Nombre From Cliente cl Inner Join ClienteTipoRuc cr On cr.idCliente = cl.idCliente Where cr.idTipoRuc = 11", true);
						dr.Read();
						if (dr.HasRows)
						{
							this.txtNumIdCliente.Value = dr.GetInt32(0);
							this.txtRuc.Value = dr.GetString(1);
							this.txtNombre.Value = dr.GetString(2);
							//this.txtClienteNombre.Enabled = true;									
						}
						dr.Close();
					}					
					#endregion Consumidor Final
				}
			}
		}
		
		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			if((int)this.cmbTipoRuc.Value == 11)
			{
				sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc1.Text.ToString(), cdsSeteoF);
				if (sParametros[0] != null) this.txtNumIdCliente1.Value = int.Parse(sParametros[0]);				
				if (sParametros[3] != null) this.txtRuc1.Text = sParametros[3];
				if (sParametros[4] != null) this.txtClienteNombre.Text = sParametros[4];
				if (sParametros[13] != null) bClienteFacturaBajoCosto = bool.Parse(sParametros[13]);
			}
			else 
			{
				sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc.Text.ToString(), cdsSeteoF);
				if (sParametros[0] != null) this.txtNumIdCliente.Value = int.Parse(sParametros[0]);
				if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
				if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
				if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];		
				if (sParametros[13] != null) bClienteFacturaBajoCosto = bool.Parse(sParametros[13]);
			}		
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{				
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 1);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}			
				//				if ((int)this.cmbFormaPago.Value == 19)
				//				{
				//					string sIdentificacion =	FuncionesProcedimientos.AbreFacturacionClientes1(false,1,(int)this.txtNumIdCliente.Value);
				//					if (sIdentificacion.Length > 0)
				//					{						
				//						this.txtRuc.Text = sIdentificacion;
				//						CargaCliente();
				//						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				//						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				//					}
				//				}
				//				else
				//				{
				//					string sIdentificacion = FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true);
				//					if (sIdentificacion.Length > 0)
				//					{						
				//						this.txtRuc.Text = sIdentificacion;
				//						CargaCliente();
				//						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				//						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				//					}
				//				}				
			}
		}

		private void txtRuc1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{				
				string sIdentificacion1 =	FuncionesProcedimientos.AbreFacturacionClientes1(false,1,(int)this.txtNumIdCliente1.Value);
				if (sIdentificacion1.Length > 0)
				{						
					this.txtRuc1.Text = sIdentificacion1;
					CargaCliente();
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}			
			}
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (e.KeyCode == Keys.F3)
				{
					if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion && !bRebates)
					{
						if (idTipoOrigen > 0)
						{
							string sDocumento = "PEDIDOS";
							if (idTipoOrigen == 11) sDocumento = "RESERVAS";
							if (idTipoOrigen == 43) sDocumento = "PREFACTURAS";

							MessageBox.Show("Este documento proviene de " + sDocumento + ", no puede ser modificado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						
						if (!FuncionesInventario.vSerialesIngresadasGridAbreInventario(this.ultraGrid1, this.grdSerial, bNuevo)) return;

						//if (!Validacion.bValidaAbreListaArticulos(this.ultraGrid2)) return;

						FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, (int)this.cmbBodega1.Value, 1, (int)this.cmbFormaPago.Value, Convert.ToDateTime(this.dtFecha.Value), false);
					}
				}
			}

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	
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
									
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 3);
			FuncionesProcedimientos.FormatoGrid(e, "Aporte", 2);
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
						if ((int)dr.Cells["Tipo"].Value != 3)	dPrecio = dPrecio + Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
					}
				}
			}
			
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			dPrecio = dPrecio * dPorcIva;

			return dPrecio;
		}

		private void Total()
		{
			if (idTipoOrigen == 3) return;

			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;
			
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;
			this.txtIva0.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtProteccionDD.Value = 0.00m;

			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100.00m, 2); //Math.Round((Convert.ToDecimal( dr.Cells["Cantidad"].Value * Convert.ToDecimal( dr.Cells["Precio"].Value) * Convert.ToDecimal( dr.Cells["DescuentoPorc"].Value/100, 2);

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

				this.txtIvaTotal.Value = Math.Round(((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva)/100.00m, 2);
			
				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 2);
			}
			#endregion Desplazamiento de grilla			

			if ((int)this.cmbFormaPago.Value == 3)
				this.txtProteccionDD.Value = FuncionesInventario.FacturacionValorDYD(this.ultraGrid1, (int)this.cmbFormaPago.Value, 0);
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
				#region Sin Clave de Desuento
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
				if ((int)this.cmbFormaPago.Value == 1 && !estadoFF.EsFF) idTPrecio = 2;
				if ((int)this.cmbFormaPago.Value == 11 && !estadoFF.EsFF) idTPrecio = 90;

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
				#endregion Sin Clave de Descuento
			}
			else
			{
				#region Con Clave de Descuento
				decimal dDescuentoTotal = dClavePorcDscto + Convert.ToDecimal(e.Cell.OriginalValue);
				
				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > dDescuentoTotal)
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("El Descuento no puede ser mayor al Permitido Descuento Original {0} + Descuento Permitido {0}", Convert.ToDecimal(e.Cell.OriginalValue), dClavePorcDscto), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
				#endregion Con Clave de Descuento
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
			string sSQLInf = string.Format("Exec CargaInformacionArticulo {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}'", 
				idArticulo, (int)this.cmbFormaPago.Value, idTipoPrecio, (int)this.cmbBodega.Value, this.txtidBanco.Value, this.txtidTarjeta.Value, this.txtidPlan.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
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
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select Inventario From Articulo Where idArticulo = {0}) = 1) Set @Res = 1 Select @Res", idArticulo), true))
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

		private void GarantiaExtendida()
		{
			string sIdArticulos = "";

			if (Funcion.bEscalarSQL(cdsSeteoF, "Select GarantiaExtendida From SeteoF", true))
			{
				#region Articulos en el Grid
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						if (sIdArticulos.Length == 0) sIdArticulos = dr.Cells["idArticulo"].Value.ToString();
						else if (sIdArticulos.Length > 0) sIdArticulos = sIdArticulos + ", " + dr.Cells["idArticulo"].Value.ToString();										
					}
				}
				#endregion Articulos en el Grid

				if (sIdArticulos.Length > 0)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaArticulosGarantiaExtendida '{0}'", sIdArticulos)) > 0)
					{
						using (frmProductosGarantiaExtendida GE = new frmProductosGarantiaExtendida(this.ultraGrid1, NumCuotas, Convert.ToDateTime(this.dtFecha.Value), (int)this.txtIdEF.Value))
						{
							if (DialogResult.OK == GE.ShowDialog())
							{
								if (GE.ultraGrid1.Rows.Count > 0)
								{
									#region Carga Productos
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in GE.ultraGrid1.Rows.All)
									{
										if ((bool)dr.Cells["Sel"].Value)
										{
											string sSQLInf = "Select a.idArticulo, a.Codigo, a.Articulo From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where GE = 1";									
											SqlDataReader drA = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
											drA.Read();									
											this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = drA.GetInt32(0);
											this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drA.GetString(1);
											this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = drA.GetString(2);
											this.ultraGrid1.ActiveRow.Cells["idUnidad"].Value = 1;
											this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value = 1;
											this.ultraGrid1.ActiveRow.Cells["Precio"].Value = dr.Cells["Precio"].Value;
											this.ultraGrid1.ActiveRow.Cells["Impuesto"].Value = dPorcIva;
											this.ultraGrid1.ActiveRow.Cells["DescuentoArt"].Value = 0.00m;
											this.ultraGrid1.ActiveRow.Cells["DescuentoPorc"].Value = 0.00m;
											this.ultraGrid1.ActiveRow.Cells["Tipo"].Value = 0;
											this.ultraGrid1.ActiveRow.Cells["Entregado"].Value = false;
											this.ultraGrid1.ActiveRow.Cells["idArticuloGE"].Value = dr.Cells["idArticulo"].Value;											
											drA.Close();						

											this.ultraGrid1.Rows.Band.AddNew();
										}
									}
									#endregion Carga Productos
								}
							}
						}
					}
				}
			}
		}
		
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bEdicion) return;
			if (bFundas)return;

			try
			{
				if (bCargaPromocionFP) return;								

				if (!bActualiza) return;
				
				Total();
				bActualiza = false;

				int VerificaCambio = 0;

				#region Codigo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					int idFormaPagoV = (int)this.cmbFormaPago.Value;
					if (idFormaPagoV == 21)
					{
						idFormaPagoV = 9;
					}
					if (idFormaPagoV == 9)
					{
						this.NumCuotas = 18;
					}
					if ((int)this.cmbFormaPago.Value == 4) idFormaPagoV = (int)this.txtNumidFPNC.Value;
					if (idFormaPagoV == 0) idFormaPagoV = (int)this.cmbFormaPago.Value;

					bActualiza = false;

					if (!bModificaColor)
						FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid1, idFormaPagoV, idTipoPrecio, (int)this.cmbBodega.Value, (int)this.cmbBodega1.Value, 
							(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, Convert.ToDateTime(this.dtFecha.Value), 1, 
							true, bPromociones, bPromociones, bPromociones, this.NumCuotas, (int)this.txtIdEF.Value, (int)this.txtNumIdCliente.Value, iNumCompras, 0, bCuponWeb, false, idCompraOrigen, cdsSeteoF, true, 0, bEmpleado);
				
					if (!bRebates && this.ultraGrid1.Rows.Count > 0) 	this.ultraGrid1.Rows.Band.AddNew();

					if (this.ultraGrid1.Rows.Count > 0) if ((int)this.cmbBodega.Value == 69) this.txtClienteNombre.Enabled = false;

					bActualiza = true;
					this.Total();	
				}		
				#endregion Codigo			

				#region Cantidad
				if (e.Cell.Column.ToString() == "Cantidad")
				{
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) 
					{
						//if (Convert.ToInt32(e.Cell.Row.Cells["idArticulo"].Value) != 12920 && Convert.ToInt32(e.Cell.Row.Cells["idArticulo"].Value) != 12921)
						if (Convert.ToInt32(e.Cell.Row.Cells["idArticulo"].Value) != 12929 && Convert.ToInt32(e.Cell.Row.Cells["idArticulo"].Value) != 12930)
						{
							bActualiza = false;
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
										bActualiza = true;
										return;
									}

									if ((int)e.Cell.Row.Cells["Tipo"].Value > 0) 
									{
										e.Cell.Value = 1;
										MessageBox.Show("Para Productos que forman parte de un Combo o Promocion la Cantidad es 1", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										bActualiza = true;
										return;
									}

									#region Existencia
									DateTime dtFecha = (DateTime)this.dtFecha.Value;
									int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
									int Existencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3})", 
										dtFecha.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

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
									#endregion Existencia
								}
							}
						}
						
					}
					bActualiza = true;
					Total();
				}
				#endregion Cantidad

				#region Precio
				if (e.Cell.Column.ToString() == "Precio")
				{				
					bActualiza = false;
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Precio"))
					{							
						Total();
						bActualiza = true;
						return;
					}
	
					bActualiza = true;

					bActualiza = false;

					int idTipoPreciovalida = idTipoPrecio;
					if ((int)this.cmbFormaPago.Value == 1 && !estadoFF.EsFF) idTipoPreciovalida = 2;
					if ((int)this.cmbFormaPago.Value == 13 && !estadoFF.EsFF) idTipoPreciovalida = 3;
					if ((int)this.cmbFormaPago.Value == 19 && !estadoFF.EsFF) idTipoPreciovalida = 3;
					if ((int)this.cmbFormaPago.Value == 11 && !estadoFF.EsFF) idTipoPreciovalida = 90;					
					if ((int)this.cmbFormaPago.Value == 3 && (int)this.txtidPlan.Value == 3 && !estadoFF.EsFF) idTipoPreciovalida = 92;
					
					if (!Facturacion.bValidaModificacionPrecios(e, miAcceso.CambiarPrecio, 1, idTipoPreciovalida, (int)this.txtNumIdCliente.Value, cdsSeteoF)) 
					{							
						Total();
						bActualiza = true;
						return;
					}

					bActualiza = true;
					Total();												
				}
				#endregion Precio

				#region IVA
				if (e.Cell.Column.ToString() == "Impuesto")
				{
					bActualiza = false;
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Impuesto"))
					{							
						Total();
						bActualiza = true;
						return;
					}

					if (!Facturacion.bValidaImpuestos(e, (DateTime)this.dtFecha.Value, 1, cdsSeteoF))
					{							
						Total();
						bActualiza = true;
						return;
					}				
						
					bActualiza = true;
				}
				#endregion IVA

				#region Descuento Porcentaje
				if (e.Cell.Column.ToString() == "DescuentoPorc")
				{		
					bActualiza = false;
											
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoPorc"))
					{							
						Total();
						bActualiza = true;
						return;
					}

					bActualiza = true;
					
					if (MenuLatinium.stIdDB == 1)
					{
						string squeryCompraUsuarioPuedeValidarDescuento = string.Format("Exec CompraUsuarioPuedeValidarDescuento {0}", MenuLatinium.IdUsuario);
						bool bCompraUsuarioPuedeValidarDescuento = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, squeryCompraUsuarioPuedeValidarDescuento, false);

						//if (MenuLatinium.IdUsuario != 1322 && MenuLatinium.IdUsuario != 345 && MenuLatinium.IdUsuario != 461 && MenuLatinium.IdUsuario != 1222)
						if (bCompraUsuarioPuedeValidarDescuento)
						{
							if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) 
							{
								if (bDescuentoActivo)
								{						
									int idTPDscto = 0;
						
									if ((int)this.cmbFormaPago.Value == 1) idTPDscto = 2;
									if ((int)this.cmbFormaPago.Value == 11) idTPDscto = 3;
									if ((int)this.cmbFormaPago.Value == 3 && (int)this.txtidPlan.Value != 3) idTPDscto = 3;
									if ((int)this.cmbFormaPago.Value == 3 && (int)this.txtidPlan.Value == 3) idTPDscto = 92;

									string sSQL = string.Format("Exec FAC_ValidaDescuentoMargen {0}, {1}, {2}", (int)e.Cell.Row.Cells["idArticulo"].Value, idTPDscto,MenuLatinium.IdUsuario);
									dClavePorcDscto = Funcion.decEscalarSQL(cdsSeteoF, sSQL);
								}
							}
						}
					}

					int idTipoPrecioValida = idTipoPrecio;
					if ((int)this.cmbFormaPago.Value == 1 && !estadoFF.EsFF) idTipoPrecioValida = 2;
					if ((int)this.cmbFormaPago.Value == 11 && !estadoFF.EsFF) idTipoPrecioValida = 90;

					if (!Facturacion.bValidaDescuentos(e, this.ultraGrid1, (int)this.cmbBodega.Value, bDescuentoActivo, dClavePorcDscto, 1, miAcceso.CambiarDescuento, idTipoPrecioValida, 
						(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, dPorcentajeCuponWeb, 0.00m, bCuponWeb, idMail, false, cdsSeteoF))
					{							
						Total();
						bActualiza = true;
						return;
					}														

					bActualiza = false;
				}

				if (e.Cell.Column.ToString() == "DescuentoArt")
				{						
					Total();
					bActualiza = true;
				}
				#endregion Descuento Porcentaje

				Total();

				bActualiza = true;
			}
			catch (Exception ex2)
			{
				MessageBox.Show(string.Format("Error al insertar filas\n\nLa Pantalla se cerrara vuelva a abrirla para continuar", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
				this.btnCancelar_Click(sender, e);
				this.Close();				
			}
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
						
			if (bEdicion && !bNuevo && (int)this.txtNumIdCompra.Value > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetCompra Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) < this.ultraGrid1.Rows.Count)
				{
					MessageBox.Show("Se eliminaron los productos de la factura vuelva a cargarlos y Guarde la Factura, para poder cancelar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}			
			}

			if (bEdicion)
			{
				if ((int)this.txtNumidOrigen.Value > 0)
				{
					string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueoOrigen, true);					
				}			
			}
			
			if (idCompraOrigen > 0 && bNuevo) this.DialogResult = DialogResult.No;
			
			bEdicion = false;
			bNuevo = false;
			
			#region Audita Descuento
			if (bDescuentoActivo)
			{
				string sMensaje = "";
					
				int iContador = 1;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (iContador > 1) sMensaje = " - "; 

					sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

					iContador++;
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'FACTURA', {1}, 4, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje));
			}
			#endregion Audita Descuento

			#region Controles
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbBodega1.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.cmbNumContadoCredito.Value = System.DBNull.Value;
			this.cmbFechaPrimerPago.Value = System.DBNull.Value;
			
			this.btnEntrega.Visible = false;
			this.btnEntrega.Enabled = false;

			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtRuc1.Text = "";
			this.txtNombre.Text = "";
			this.txtComprobante.Text = "";
			this.lblEstado.Text = "";
			this.txtClienteNombre.Text = "";

			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today.AddDays(-365);
			this.dtFechaRetencion.Value = System.DBNull.Value;

			this.txtNumEstado.Value = 0;
			this.txtNumIdAsiento.Value = 0;
			this.txtNumIdCliente.Value = 0;
			this.txtNumIdCliente1.Value = 0;
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
			this.txtProyecto.Text = "";
			this.lblUsuario.Text = "";
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtBaseImponible.Value = 0.00m;
			this.optTipoVenta.EditValue = 0;
			this.txtresponsablecurier.Text = "";
			this.txtNumGuia.Text = "";
			this.txtValorEntrega.Value = 0.00m;
			this.txtNumTransaccion.Text = "";
			this.txtDireccionEntrega.Text = "";
			this.lblYupana.Text = "";
			this.dtFechaEntregaCuriermbo1.Value = System.DBNull.Value;
			this.cmbCourier.Enabled = true;
			this.chkEntregaLocal.Checked = false;
			this.chkEntregaLocal.Enabled = false;
			this.chkFreelance.Checked = false;
			this.chkFreelance.Visible = false;
			this.chkEntregaLocal.Enabled = false;
			this.btnEntrega.Visible = false;
			this.btnEntrega.Enabled = false;

			dCuotaAsignada = 0.00m;
			dInteres = 0.00m;
			dCuotaDisponible = 0.00m;
			dCuotaUsada = 0.00m;
			AutorizaEntrada = false;
			dValorCuota = 0.00m;

			this.cmbFormaPago.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbBodega1.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.cmbNumContadoCredito.Enabled = false;
			this.cmbFechaPrimerPago.Enabled = false;


			this.txtNumero.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtRuc1.Enabled = false;
			this.txtClienteNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaRetencion.Enabled = false;
						
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnCancelar.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimeED.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnTemporal.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Text = "Anular";
			this.btnCliente.Enabled = false;
			this.btnCliente2.Enabled = false;
			this.btnValidarMail.Enabled = false;

			this.lblPlan.Text = "";

			this.txtresponsablecurier.Enabled = true;
			this.txtNumGuia.Enabled = true;
			this.txtValorEntrega.Enabled = true;
			this.txtNumTransaccion.Enabled = true;
			this.txtDireccionEntrega.Enabled = true;
			this.dtFechaEntregaCuriermbo1.Enabled = true;
			this.cmbCourier.Enabled = true;

			this.mnuImprimirFactura.Enabled = false;
			this.mnuImprimirContrato.Enabled =false;
			this.mnuImprimirNuovo.Enabled =false;
			this.mnuImprimirGarantiaExtendida.Enabled = false;
			this.mnuDocsCrediPoint.Enabled = false;
			this.mnuDocsCrediPointBlanco.Enabled = false;
			this.mnuImprimirDYD.Enabled = false;
			this.mnuImprimirRIDE.Enabled = false;

			if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;

			this.btnClaveDescuento.Enabled = false;

			HabilitaTab(false);

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);			
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");

			if (idCuponWeb > 0) 
				Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, (int)this.txtNumIdCompra.Value, cdsSeteoF);
			#endregion Controles

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
			this.chkEntregaLocal.Checked = false;
			this.cmbDirecciones.Value = System.DBNull.Value;
			this.dtFechaEntrega.Value = System.DBNull.Value;
			this.dtpHoraEntrega.Value = System.DBNull.Value;
			#endregion Tab Entrega
			#region Tab Nota
			this.txtNotas.Text = "";	
			#endregion Tab Nota
			#endregion Tab

			#region Variables
			iAnticipoCDP = 0;
			iAsientoCDP = 0;
			bFormaPagoPromocional = false;
			bCargaPromocionFP = false;
			idTipoPrecio = 0;
			bNumeracionAutomatica = false;
			bPromociones = false;
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
			idBloqueaTransacciones1 = 0;
			dtFecVencCupon = DateTime.Today;
			sNumero = "";
			dClavePorcDscto = 0.00m;
			bVeriricador = false;
			dPorcIva = 0.00m;
			idCuponWeb = 0;
			bValidaTipoVenta = false;
			iNumCompras = 0;
			iBodegaDestino = 0;
			bCuponWeb = false;
			idMail = 0;
			idFe_EstadoDocumento = 0;
			iEncuesta = 0;
			sOtrosEncuesta = "";
			CodigoEntrega ="";
			SettearFamilyFriends();
			idCre_SolicitudWeb = 0;
			#endregion Variables

			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");
			
			if (bServicioTecnico) this.DialogResult = DialogResult.No;		
			
			#region FamilyFriends
			SettearFamilyFriends();
			#endregion FamilyFriends
		}

		private void SettearFamilyFriends()
		{
			reqFF = new FamilyFriendsRequestModel();
			this.lblRecurrente.Text = "";
			estadoFF = new FamilyFriendsEstadoModel();
		}

		private void ImpresionGarantiaExtendidaPlus()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				#region Imprime reporte de Garantia Extendida
				if ((int)dr.Cells["Tipo"].Value == 9)
				{								
					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramFieldIdDetCompra = new ParameterField ();
					ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
					paramFieldIdDetCompra.ParameterFieldName = "@idDetCompra";
					discreteValIdDetCompra.Value = (int)dr.Cells["idDetCompra"].Value; 
					paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
					paramFields.Add (paramFieldIdDetCompra);
									
					Reporte Reporte = new Reporte("GarantiaExtendidaPlus.rpt", "");
					Reporte.MdiParent = this.MdiParent;
					Reporte.Titulo("Garantia Extendida");
					Reporte.crVista.ParameterFieldInfo = paramFields;
					Reporte.Show();	
					
					int idArticuloGEx = (int)dr.Cells["idArticuloGE"].Value;
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drArt in this.ultraGrid1.Rows.All)
					{
						if (idArticuloGEx == (int)drArt.Cells["idArticulo"].Value)
						{
							ParameterFields paramFieldsCP = new ParameterFields ();
							ParameterField paramFieldIdDetCompraCP = new ParameterField ();
							ParameterDiscreteValue discreteValIdDetCompraCP = new ParameterDiscreteValue ();
							paramFieldIdDetCompraCP.ParameterFieldName = "@idDetCompra";
							discreteValIdDetCompraCP.Value = (int)drArt.Cells["idDetCompra"].Value; 
							paramFieldIdDetCompraCP.CurrentValues.Add (discreteValIdDetCompraCP);
							paramFieldsCP.Add (paramFieldIdDetCompraCP);
								
							ParameterField paramFieldIdDetCompraCPT = new ParameterField ();
							ParameterDiscreteValue discreteValIdDetCompraCPT = new ParameterDiscreteValue ();
							paramFieldIdDetCompraCPT.ParameterFieldName = "@Tipo";
							discreteValIdDetCompraCPT.Value = (int)drArt.Cells["Tipo"].Value; 
							paramFieldIdDetCompraCPT.CurrentValues.Add (discreteValIdDetCompraCPT);
							paramFieldsCP.Add (paramFieldIdDetCompraCPT);
								
							Reporte ReporteCP = new Reporte("CompraProtegida_1.rpt", "");
							ReporteCP.MdiParent = this.MdiParent;
							ReporteCP.Titulo("Compra Protegida");
							ReporteCP.crVista.ParameterFieldInfo = paramFieldsCP;
							ReporteCP.Show();
						}
					}					
				}
				#endregion Imprime reporte de Garantia Extendida
			}
			//			ImpresionCompraProtegida();
		}

		private void ImpresionGarantiaExtendida()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				#region Imprime reporte de Garantia Extendida
				if ((int)dr.Cells["Tipo"].Value == 5)
				{								
					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramFieldIdDetCompra = new ParameterField ();
					ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
					paramFieldIdDetCompra.ParameterFieldName = "@idDetCompra";
					discreteValIdDetCompra.Value = (int)dr.Cells["idDetCompra"].Value; 
					paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
					paramFields.Add (paramFieldIdDetCompra);
									
					Reporte Reporte = new Reporte("GarantiaExtendidaPlus.rpt", "");
					Reporte.MdiParent = this.MdiParent;
					Reporte.Titulo("Garantia Extendida");
					Reporte.crVista.ParameterFieldInfo = paramFields;
					Reporte.Show();						
				}
				#endregion Imprime reporte de Garantia Extendida
			}
		}

		private void ImpresionCompraProtegida()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				#region Imprime reporte de Compra Protegida
				if ((int)dr.Cells["Tipo"].Value == 10 ||(int)dr.Cells["Tipo"].Value == 1)
				{	
					int idArticuloGEx = (int)dr.Cells["idArticuloGE"].Value;
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drArt in this.ultraGrid1.Rows.All)
					{
						if (idArticuloGEx == (int)drArt.Cells["idArticulo"].Value)
						{		
							ParameterFields paramFields = new ParameterFields ();
							ParameterField paramFieldIdDetCompra = new ParameterField ();
							ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
							paramFieldIdDetCompra.ParameterFieldName = "@idDetCompra";
							discreteValIdDetCompra.Value = (int)drArt.Cells["idDetCompra"].Value; 
							paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
							paramFields.Add (paramFieldIdDetCompra);

							ParameterField paramFieldIdDetCompraCPT = new ParameterField ();
							ParameterDiscreteValue discreteValIdDetCompraCPT = new ParameterDiscreteValue ();
							paramFieldIdDetCompraCPT.ParameterFieldName = "@Tipo";
							discreteValIdDetCompraCPT.Value = (int)dr.Cells["Tipo"].Value;; 
							paramFieldIdDetCompraCPT.CurrentValues.Add (discreteValIdDetCompraCPT);
							paramFields.Add (paramFieldIdDetCompraCPT);
									
							Reporte Reporte = new Reporte("CompraProtegida_1.rpt", "");
							Reporte.MdiParent = this.MdiParent;
							Reporte.Titulo("Compra Protegida");
							Reporte.crVista.ParameterFieldInfo = paramFields;
							Reporte.Show();	
							continue;
						}
					}
				}
				#endregion Imprime reporte de Compra Protegida				
			}
		}


		private void ImpresionPointRespaldo()
		{
			int idCompra = (int)this.txtNumIdCompra.Value;

			// Crear parámetros para el reporte
			ParameterFields paramFields = new ParameterFields();
			ParameterField paramFieldIdCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdCompra = new ParameterDiscreteValue ();
			paramFieldIdCompra.ParameterFieldName = "@idCompra";
			discreteValIdCompra.Value = idCompra; 
			paramFieldIdCompra.CurrentValues.Add (discreteValIdCompra);
			paramFields.Add (paramFieldIdCompra);

			// Cargar y mostrar el reporte
			Reporte reporte = new Reporte("PointRespaldo.rpt", "");
			reporte.MdiParent = this.MdiParent;
			reporte.Titulo("Point Respaldo");
			reporte.crVista.ParameterFieldInfo = paramFields;
			reporte.Show();
		}

		private void ImpresionAsistenciaSaludDental()
		{
			int idCompra = (int)this.txtNumIdCompra.Value;

			// Crear parámetros para el reporte
			ParameterFields paramFields = new ParameterFields();
			ParameterField paramFieldIdCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdCompra = new ParameterDiscreteValue ();
			paramFieldIdCompra.ParameterFieldName = "@idCompra";
			discreteValIdCompra.Value = idCompra; 
			paramFieldIdCompra.CurrentValues.Add (discreteValIdCompra);
			paramFields.Add (paramFieldIdCompra);

			// Cargar y mostrar el reporte
			Reporte reporte = new Reporte("AsistenciaSaludDental.rpt", "");
			reporte.MdiParent = this.MdiParent;
			reporte.Titulo("Asistencia Salud Dental");
			reporte.crVista.ParameterFieldInfo = paramFields;
			reporte.Show();
		}

		
		private void ImpresionProductoDyD()
		{
			int iCont = 0;

			//int idArt = Funcion.iEscalarSQL(cdsSeteoF, "Select a.idArticulo From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where Seguro = 1 And Discontinuo = 0");

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				if ((int)dr.Cells["idArticulo"].Value == 8683) iCont++;
  	
			//string sReporte = "ProteccionDYD.rpt";
			//if (idArt == 8683)
			string sReporte = "DatosSeguroDesempleo.rpt";

			if (iCont > 0)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = (int)this.txtNumIdCompra.Value; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte(sReporte, "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Proteccion D&D");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.ShowDialog();				
			}
		}

		private void ImpresionDirecta(int iGarantias)
		{
			if (bNuevo)
			{
				using (frmMensajeNumeración Msje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, iGarantias, false))
				{
					if (DialogResult.OK == Msje.ShowDialog())
					{
						Imprimir(1);
						
						ImpresionGarantiaExtendida();

						ImpresionProductoDyD();
					}
				}
			}
		}

		public string Modulo11(string number)
		{
			int pivote = 2;
			int longitudCadena = number.Length;
			int cantidadTotal = 0;
			int b = 1;

			for (int i = 0; i < longitudCadena; i++)
			{
				if (pivote == 8) pivote = 2;

				int temporal = int.Parse(number.Substring(i, b));

				temporal *= pivote;

				pivote++;

				cantidadTotal += temporal;
			}

			cantidadTotal = 11 - cantidadTotal % 11;

			if (cantidadTotal == 11) cantidadTotal = 0;
			else if (cantidadTotal == 10) cantidadTotal = 1;

			return cantidadTotal.ToString();
		}

		public string invertirCadena(String cadena)
		{
			String cadenaInvertida = "";
			for (int x = cadena.Length - 1; x >= 0; x--)
				cadenaInvertida = cadenaInvertida + cadena.Substring(x, 1);

			return cadenaInvertida;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			string sSQL = "";
			decimal dRetenido = 0;
			decimal dRetenidoIVA = 0;
			decimal dRetenidoRenta = 0;
			int idDetFormaPagoRetencion = 0;
			DateTime dtFechaCadR = DateTime.Today;
			int idProformaAseguradora = 0;

			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "VENTA", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (!bRebates && bExigeVendedor)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;
			
			if (bNuevo)
				if (!Validacion.vbNumeracion(bNumeracionAutomatica, (int)this.txtNumIdCompra.Value, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
			
			if (MenuLatinium.stIdDB != 38)
			{
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;

				if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbFormaPago.Value, cdsSeteoF)) return;		
			}		
			#region entregaDomiciliogps
			/*
			if (  (bool) this.chkEntregaLocal.Checked == false || (bool) this.chkEntregado.Checked == false )
			{
				MessageBox.Show("Debe seleccionar si la entrega será en domicilio o en local.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((bool) this.chkEntregado.Checked )
			{
				if( CodigoEntrega.Length == 0 )
				{
					MessageBox.Show("Seleccione la ubicación de entrega antes de continuar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				int CountGPS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaGeoreferenciaEntrega] '{0}','{1}',{2}",
					CodigoEntrega, this.txtRuc.Text, (int) this.cmbBodega.Value ));	

				if( CountGPS = 0)
				{
					MessageBox.Show("Seleccione la ubicación de entrega antes de continuar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if( CountGPS > 1)
				{
					MessageBox.Show("Existe mas de una ubicacion comunicate con el Area Desarrollo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}*/
			#endregion entregaDomiciliogps
			if (!Validacion.EntregaDomicilio(this.chkEntregado, this.cmbDirecciones, this.dtFechaEntrega, this.dtpHoraEntrega, (int)this.optTipoVenta.EditValue)) return;

			if (!Validacion.Curier(this.chkEntregado, this.cmbCourier, this.dtFechaEntregaCuriermbo1,this.txtNumGuia,(int)this.optTipoVenta.EditValue)) return;
			
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 1)) return;	
	
			if((int)this.cmbFormaPago.Value == 19)
				if (!Validacion.vbComboVacio(this.cmbCourier, "Por la forma de pago (Contraentrega) debe ingresar los datos del courier")) return;

			if (MenuLatinium.stIdDB != 38)
			{
				if((int)this.cmbTipoRuc.Value == 11)
				{
					if (Convert.ToDecimal(this.txtTotal.Value) > 200.00m)
					{
						MessageBox.Show("Pasado los 200 dolares debe ingresar de forma obligatoria los datos del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
			}

			#region Valida Promocion sin intereses
			//			if ((int)this.cmbFormaPago.Value == 3)
			//			{
			//				if ((int)this.txtidPlan.Value == 3)
			//				{		
			//					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*) FROM Bodega WHERE {0} NOT IN (29, 31, 85, 38, 16, 96, 15, 84, 13)", 
			//						(int)this.cmbBodega.Value)) > 0)
			//					{
			//						MessageBox.Show("No tiene autorización para facturar sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}
			//
			//					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Banco Where {0} NOT IN (6, 7)", 
			//						(int)this.txtidBanco.Value)) > 0)
			//					{
			//						MessageBox.Show("El banco seleccionado no tiene autorización para facturar sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}
			//				
			//					int iContProducto = 0;
			//					int iContPromociones = 0;
			//					int iContDescuento = 0;
			//					int iContOtros = 0;
			//
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			//					{
			//						string sSQLQS = string.Format("Select COUNT(*) From Articulo Where {0} In (6365, 8268, 6621, 6619, 8267, 6394, 6625)", 
			//							(int)dr.Cells["idArticulo"].Value);
			//						if (Funcion.iEscalarSQL(cdsSeteoF, sSQLQS) > 0)
			//						{
			//							iContProducto++;
			//
			//							if (Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) > 0.00m) iContDescuento++;
			//						}
			//
			//						if ((int)dr.Cells["Tipo"].Value != 0)
			//							if ((int)dr.Cells["Tipo"].Value != 5)
			//								if ((int)dr.Cells["Tipo"].Value != 3) 
			//									iContPromociones++;
			//
			//						sSQLQS = string.Format("Select COUNT(*) From Articulo Where {0} Not In (6365, 8268, 6621, 6619, 8267, 6394, 6625)", (int)dr.Cells["idArticulo"].Value);
			//						if (Funcion.iEscalarSQL(cdsSeteoF, sSQLQS) > 0)							
			//							if ((int)dr.Cells["Tipo"].Value != 5)
			//								if ((int)dr.Cells["Tipo"].Value != 3) 
			//									iContOtros++;
			//					}
			//
			//					if (iContProducto == 0)
			//					{
			//						MessageBox.Show("Los productos cargados no estan autorizados a facturar sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}
			//
			//					if (iContDescuento > 0)
			//					{
			//						MessageBox.Show("No estan autorizados descuentos para facturar sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}
			//
			//					if (iContPromociones > 0)
			//					{
			//						MessageBox.Show("No esta autorizado a facturar combos ni promociones adicionales sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}
			//
			//					if (iContOtros > 0)
			//					{
			//						MessageBox.Show("Los productos ingresados no estan autorizados a facturar sin intereses con pagos en tarjeta de crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//						return;
			//					}												
			//				}
			//			}
			#endregion Valida Promocion sin intereses

			#region Valida Promocion QS Bloqueada
			//			bool bValidaPromObligatorio = false;
			//			int iCont2 = 0;
			//			int iCont1 = 0;
			//
			//			if ((int)this.cmbBodega.Value == 29)
			//			{
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			//				{
			//					string sSQLQS = string.Format("Select COUNT(*) From ArticuloSubGrupo s Inner Join Articulo a On a.idSubGrupo = s.idSubGrupo Where a.idSubGrupo In (66, 131, 189, 203, 205, 220) And a.idArticulo = {0}", 
			//						(int)dr.Cells["idArticulo"].Value);
			//					if (Funcion.iEscalarSQL(cdsSeteoF, sSQLQS) > 0) iCont1++;	
			//				}
			//
			//				if (iCont1 == 0)
			//				{
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			//					{
			//						string sSQLQS = string.Format("Select COUNT(*) From ArticuloSubGrupo s Inner Join Articulo a On a.idSubGrupo = s.idSubGrupo Where a.idSubGrupo In (146, 105) And a.idArticulo = {0}", 
			//							(int)dr.Cells["idArticulo"].Value);
			//						if (Funcion.iEscalarSQL(cdsSeteoF, sSQLQS) > 0) iCont2++;	
			//					}
			//				}
			//
			//				if (iCont2 > 0) bValidaPromObligatorio = true;
			//			}
			#endregion Valida Promocion QS Bloqueada

			#region Valida Promocion Producto Obligatorio 
			//			if (!bValidaPromObligatorio)
			//			{
			int idPromProdObl = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionProductoObligatorio({0}, {1}, {2}, {3}, {4}, {5}, {6}, 0, 0)",
				(int)this.cmbBodega.Value, 0, idTipoPrecio, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value));
			if (idPromProdObl > 0)
			{
				int idArticuloPO = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idArticulo From DetallePromocion Where idPromocion = {0}", idPromProdObl));
				int iContPO = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (idArticuloPO == (int)dr.Cells["idArticulo"].Value) iContPO++;					
				}

				if (iContPO == 0)
				{
					MessageBox.Show("Existe promocion por la forma de pago seleccionada y no ha registrado el producto promocional", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			//			}
			#endregion Valida Promocion Producto Obligatorio
				
			#region Valida Existencia
			if (!bConsignacion) 
				if (idTipoOrigen == 43 || idTipoOrigen == 0)
					if (!Validacion.bValidaExistencias(this.ultraGrid1, (DateTime)this.dtFecha.Value, (int)this.txtNumIdCompra.Value, (int)this.cmbBodega1.Value, 1, idCompraOrigen, cdsSeteoF)) return;
			#endregion Valida Existencia

			/* MARGEN */
			int idTipoPrecioVM = 2;
			if ((int)this.cmbFormaPago.Value == 3) idTipoPrecioVM = 3;
			decimal dPrecioOriginal = (decimal)this.txtIva0.Value + (decimal)this.txtIva.Value;
			if (!Validacion.bValidaMargen(this.ultraGrid1, idTipoPrecioVM, MenuLatinium.stUsuario, (decimal)this.txtBaseImponible.Value, dPrecioOriginal, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, (int)this.txtNumIdCompra.Value, 10, cdsSeteoF)) return;

			if (idTipoOrigen != 3)	Total();
									
			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			#region Verifica Seriales
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega1.Value), true))
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

			#region Valida Existencia Inventario enero
		
			string sVSerial = "";
			int iVSerialVal = 0;
			if (MenuLatinium.stIdDB == 1) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
				{
					sVSerial = "";
					iVSerialVal = 0;
					if (dr.Cells["idSerial"].Value != DBNull.Value)
					{
						iVSerialVal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaSerialesConteo] {0}", (int) dr.Cells["idSerial"].Value));	
						if (iVSerialVal <= 0)
						{
							sVSerial = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaSerialesConteoMensaje] {0}", (int) dr.Cells["idSerial"].Value));

							MessageBox.Show("" + sVSerial + "\n\n  No existe por favor comunicarse con el área de \n\n Auditoria Ext: 2011/2010 ", "Point Technology");
							return;
						}
					}
				}
			}
			#endregion Valida Existencia Invenatrio enero

			#region Valida Estado Origen Nota De Credito
			if ((int)this.cmbFormaPago.Value == 4)
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.txtNumidOrigen.Value), true)) 
				{
					MessageBox.Show("La Nota De Credito Del Cliente se encuentra Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
					return;
				}
			}
			#endregion Valida Estado Origen Nota De Credito

			int idWeb = 0;
			idWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select count (*) from web_cabecera_pedido where idEstado_WEB > 2 and numero_pedido = {0}", 
				this.txtNumidOrigen.Value));
			
			if ((int)this.cmbFormaPago.Value == 32)
			{
				if (sUrlNube.Length == 0)
				{	
					MessageBox.Show("En forma de pago de una debe subir el comprobante de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			#region Valida Compra Protegida
			int iCP = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idArticulo"].Value == 8723) iCP++;
				if ((int)dr.Cells["idArticulo"].Value == 8724) iCP++;
			}

			if (iCP > 0)
			{
				sSQL = string.Format("Exec FacturacionValidaComprasProtegida {0}", (int)this.txtNumIdCliente.Value);
				int iComprasProtegidas = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
		
				if (iComprasProtegidas > 0)
				{
					MessageBox.Show("El cliente tiene compras protegidas vigentes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida Compra Protegida
			#endregion Validacion

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
			int idVendedor = 0;
			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;
			
			if ((int)this.txtNumidOrigen.Value > 0)
				FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), false, miAcceso.EditarCuentaNombre);

			#region Encuesta Anterior
			//			int iEncuesta = 0;
			//			string sOtrosEncuesta = "";
			//
			//			if (bNuevo)
			//			{
			//				if (Funcion.bEscalarSQL(cdsSeteoF, "Select Encuesta From SeteoF", true))
			//				{
			//					using (frmFacturacionEncuesta FE = new frmFacturacionEncuesta())
			//					{
			//						if (DialogResult.OK == FE.ShowDialog())
			//						{
			//							iEncuesta = (int)FE.optEncuesta.Value;
			//							sOtrosEncuesta = FE.txtOtros.Text.ToString();
			//						}
			//						else return;
			//					}
			//				}
			//			}
			#endregion Encuesta Anterior

			#region Genera factura Proforma
			if (idTipoOrigen == 11 &&	(int) this.txtNumIdCompra.Value == 0)
			{	
				int idDireccion = 0;
				int idCourier = 0;
				if (this.cmbDirecciones.ActiveRow != null) idDireccion = (int)this.cmbDirecciones.Value;
				if (this.cmbCourier.ActiveRow != null) idCourier = (int)this.cmbCourier.Value;
				bool bSeriales = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);

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
				if (this.dtpHoraEntrega.Value != System.DBNull.Value)
				{
					bHorEnt = true;
					dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
				}
				if (bCrediPoint) bFecPrimPag = true;
				#endregion Fechas
					
				#region Datos cliente

				string SRucCliente  = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select (select Ruc from Compra where idCompra = c.idorigen) from Compra c where idCompra = {0}",(int)this.txtNumidOrigen.Value));		
				int idOrigen1  = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select ISNULL (idCompra,0) from Compra c where idCompra = {0}",(int)this.txtNumidOrigen.Value));	
				int idSubProyecto  = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select ISNULL (idSubProyecto, 0) from Compra c where idCompra = {0}",(int)this.txtNumidOrigen.Value));	
        
				int iNumIdCliente = 0;
				int	iTipoRuc = 0;
				string	sRuc ="";
				string		sNombre ="";
				decimal dValor = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select (( (Precio - DescuentoArt)  * 30 )/ 100 )From Detcompra  Where idDetCompra = {0} ",idSubProyecto));		
       
				string sCodigo = "";
				int idArticulo = 0;
				string sArticulo = "";
				decimal dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

				decimal IvaTotal = Math.Round(((dValor - 0) * dPorcIva)/100.00m, 2);
			
		

				string[] sParametros = new string[11];

				sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, SRucCliente, cdsSeteoF);
				if (sParametros[0] != null) iNumIdCliente= int.Parse(sParametros[0]);
				if (sParametros[1] != null) iTipoRuc = int.Parse(sParametros[1]);
				if (sParametros[3] != null) sRuc = sParametros[3];
				if (sParametros[4] != null) sNombre = sParametros[4];	
	
				#endregion Datos cliente

				/* 9316 point*/ 
				/*9242  desarrollo*/
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select a.idArticulo, a.Codigo, a.Articulo From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where a.idArticulo = 9316", true);
				dr.Read();
				if (dr.HasRows)
				{										
					sCodigo = dr.GetString(1);
					idArticulo = dr.GetInt32(0);
					sArticulo = dr.GetString(2);
				}
				dr.Close();

				#region Graba Factura 

				string sSQLPro = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, '{46}', {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, '{62}', '{63}', {64}, '{65}', {66}, {67}, '{68}', {69}", 
					idProformaAseguradora, /*0*/
					(int)this.txtNumidOrigen.Value, /*1*/
					(int)this.txtIdSucursal.Value, /*2*/
					1, /*3*/
					(int)this.txtIdComprobante.Value, /*4*/
					6, /*5*/
					bConsignacion, /*6*/
					(int)this.txtNumContadoCredito.Value, /*7*/
					1,/*8*/
					(int)this.cmbBodega.Value, /*9*/
					(int)this.txtNumIdProyecto.Value, /*10*/
					idVendedor, /*11*/
					this.txtNumero.Text.ToString(), /*12*/
					this.txtComprobante.Text.ToString(),/*13*/
					iTipoRuc, /*14*/
					iNumIdCliente, /*15*/
					sRuc, /*16*/
					sNombre,/*17*/
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),/*18*/
					this.txtSerie.Text.ToString(), /*19*/
					this.txtAutFactura.Text.ToString(), /*20*/
					dtFechaCad.ToString("yyyyMMdd"), /*21*/
					(bool)this.chkEntregado.Checked, /*22*/
					dtFechaEnt.ToString("yyyyMMdd"), /*23*/
					dtHoraEntrega.ToString("yyyyMMdd"), /*24*/
					idDireccion, /*25*/
					this.txtNotas.Text.ToString(),/*26*/
					bCrediPoint, /*27*/
					dTotalCuotas, /*28*/
					NumCuotas, /*29*/
					Convert.ToDecimal(this.txtCuotaDeEntrada.Value), /*30*/
					dtFechPrimPago.ToString("yyyyMMdd"),/*31*/
					(decimal)this.txtIva0.Value,/*32*/
					(decimal)this.txtDescIva0Total.Value, /*33*/
					0.00m,
					//				(decimal)this.txtIva.Value, /*34*/
					(decimal)this.txtDescIvaTotal.Value, /*35*/
					0.00m, /*36*/
					Math.Round((dValor +  IvaTotal) ,2),/*2371*/
					//					dValor,
					bFecPrimPag, /*38*/
					bFecCad,/*39*/
					bFecEnt, /*40*/
					bHorEnt,/*41*/
					bAnulado, /*42*/
					true, /*43*/
					(int)this.cmbBodega1.Value,/*44*/
					0, /*45*/
					"", /*46*/
					0.00m,/*47*/
					dInteres,/*48*/ 
					dCuotaAsignada,/*49*/ 
					dCuotaDisponible,/*50*/ 
					dCuotaUsada,/*51*/
					false, /*52*/
					dValorCuota, /*53*/
					(int)this.txtIdEF.Value, /*54*/
					(int)this.txtidBanco.Value,/*55*/
					(int)this.txtidTarjeta.Value, /*56*/
					(int)this.txtidPlan.Value, /*57*/
					(int)this.txtidPlazo.Value,/*58*/
					"",/*59*/
					bVeriricador, /*60*/
					3,/*61*/
					this.txtNumTransaccion.Text.ToString(), /*62*/
					this.txtDireccionEntrega.Text.ToString(),/*63*/
					idCourier, /*64*/
					this.txtNumGuia.Text.ToString(), /*65*/
					Convert.ToDecimal(this.txtValorEntrega.Value),/*66*/
					(int)this.txtNumIdproyectoZoom.Value, /*67*/
					this.txtClienteNombre.Text.ToString(), /*68*/
					0.00m/*69*/
					);	
				
					
				idProformaAseguradora =Funcion.iEscalarSQL(cdsSeteoF,sSQLPro);

				string ssSQLi = string.Format("Update Compra Set idOrigen1 = {1} Where idCompra = {0}", idProformaAseguradora, idOrigen1);
				Funcion.EjecutaSQL(cdsSeteoF, ssSQLi);

				string sSQLLMenterga = string.Format("Update compra Set EntregadoLocal = {0}, CodigoEntrega = '{1}' Where idcompra = {2}", Convert.ToInt32( this.chkEntregaLocal.Checked) , CodigoEntrega,(int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLMenterga);

				//				#region Detalle
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drg in this.ultraGrid1.Rows.All)
				{
					int iPosicion = 0;
				

					string sSQLDetallePro = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}",
						0,
						0, 
						idProformaAseguradora,
						idArticulo, 
						1, 
						sCodigo, 
						sArticulo,
						1, 
						dValor, 
						dPorcIva, 
						0.00m, 
						0.00m, 
						0.00m,
						false, 
						8, 
						false, 
						1, 
						false, 
						false, 
						0, 
						0, 
						0.00m);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetallePro);
				}
				#endregion Graba Factura 
			}
			#endregion Genera factura Proforma

			#region Valida Anticipo Cliente
			decimal dValorcuota  = 0.00m;
					
			if (bNuevo)
			{
				if ((int)this.txtNumIdCompra.Value > 0)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Anticipos Where idOrigen = {0}", (int)this.txtNumIdCompra.Value)) == 0)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AbonaCuota From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true))
						{
							dValorcuota = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ValorCuotas From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));

							if ((int)this.cmbFormaPago.Value == 9)
							{
								using (frmAnticiposClientes Anticipos = new frmAnticiposClientes (true, this.txtRuc.Text.ToString(), (int)this.cmbBodega.Value, dValorCuota, 12, (int)this.txtNumIdCompra.Value, false))
								{
									if (DialogResult.OK == Anticipos.ShowDialog())
									{
										iAnticipoCDP = (int)Anticipos.txtIdAnticipo.Value;
										iAsientoCDP = (int)Anticipos.txtIdAsiento.Value;
									}
									else return;
								}
							}
						}
					}
				}
			}
			#endregion Valida Anticipo Cliente
			
			#region Fundas
			if (bNuevo)
			{
				if ((int)this.cmbFormaPago.Value != 9)
				{
					if (Convert.ToDecimal(this.txtTotal.Value) > 99.99m)
					{
						int iSumaCantidad=0;
						int iFundaMediana= 10717;
						int iFundaGrande= 10718;
							
						#region Existencia Fundas Medianas
						DateTime dtFechaFM = (DateTime)this.dtFecha.Value;
						int ExistenciaFundasM = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3},0,0)", 
							dtFechaFM.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, iFundaMediana));
						#endregion Existencia Fundas Medianas

						#region Existencia Fundas Grandes
						DateTime dtFechaFG = (DateTime)this.dtFecha.Value;
						int ExistenciaFundasG = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3},0,0)", 
							dtFechaFG.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, iFundaGrande));
						#endregion Existencia Fundas Grandes

						iSumaCantidad = ExistenciaFundasM + ExistenciaFundasG;

						if (iSumaCantidad > 0) Fundas();

						iSumaCantidad=0;
					}
				}
			}
			#endregion Fundas

			Total();

			#region Grabar
			#region Credito - CrediPOINT SIN ENTRADA - PROMOCIONALES
			/*valida si viene de factura proveedor*/
			int idtipoFacturaPro = 0;
			//			if ((int)this.txtNumidOrigen.Value > 0 && (int)this.cmbFormaPago.Value != 4)
			//			{
			//				idtipoFacturaPro = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select (select ISNULL (idTipoFactura,0) from Compra where idCompra = c.idOrigen) from Compra c where c.idCompra = {0} ", (int)this.txtNumidOrigen.Value));
			//			}
			/*valida si viene de factura proveedor*/
			if ((bCrediPoint && (decimal)this.txtCuotaDeEntrada.Value == 0.00m) || ((int)this.cmbFormaPago.Value == 6 && !bUsoAnticipo) || bFormaPagoPromocional)// || idtipoFacturaPro == 11)
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
				
				this.txtNumEstado.Value = 10;

				//				using (VentaFormaDePago miCobro = new VentaFormaDePago ((int)this.txtNumIdCompra.Value, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, this.txtComprobante.Text.ToString(), BUsoNotaDeCredito, idCompraOrigen, dSaldoNotaDeCredito, bRetencionCaja, dValorRetencion, bUsoAnticipo, dAnticipo, bCrediPoint, dValorCrediPoint, (decimal)this.txtCuotaDeEntrada.Value, 1, (int)this.txtNumIdCliente.Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, (int)this.cmbBodega.Value, this.ultraGrid1, (DateTime)this.dtFecha.Value, false, 0))
				//				{
				//					if (DialogResult.OK == miCobro.ShowDialog())
				//					{

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
						#region Proceso
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						#region Ingreso Egreso De Bodega Consignacion
						if (MenuLatinium.bConsignacionVenta)
						{
							if (bNuevo)
							{
								int iCont = 0;
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
									if ((bool)dr.Cells["Consignacion"].Value) iCont++;

								if (iCont > 0)
								{
									int idCompraE = 0; 
									int idCompraI = 0;
									string sNumeroIB = ""; 
									string sNumeroEB = "";
									int iBodegaCS = 0;

									oCmdActualiza.CommandText = string.Format("Select Bodega From AsignaBodegasLocales Where Local = {0} And Consignacion = 1", (int)this.cmbBodega.Value);
									iBodegaCS = (int)oCmdActualiza.ExecuteScalar();
								
									#region Numeracion
									if (bNuevo)
									{									
										oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 8, {0}, 0", iBodegaCS);
										sNumeroEB = oCmdActualiza.ExecuteScalar().ToString();

										oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 9, {0}, 0", (int)this.cmbBodega1.Value);
										sNumeroIB = oCmdActualiza.ExecuteScalar().ToString();
									}
									#endregion Numeracion

									#region Graba Registro 
									oCmdActualiza.CommandText = string.Format("Exec GrabaIEBodega {0}, 8, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, '{8}', '', 7",
										idCompraE, 171570, iBodegaCS, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumeroEB, 3, 0.00m, 0.00m, this.txtNotas.Text.ToString());;
									idCompraE = (int)oCmdActualiza.ExecuteScalar();

									oCmdActualiza.CommandText = string.Format("Exec GrabaIEBodega {0}, 9, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, '{8}', '', 7",
										idCompraI, 171570, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumeroIB, 3, 0.00m, 0.00m, this.txtNotas.Text.ToString());;
									idCompraI = (int)oCmdActualiza.ExecuteScalar();
													
									int iPosicionArticulo = 0;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
									{
										if ((bool)dr.Cells["Consignacion"].Value)
										{
											int idDetCompraE = 0;
											int idDetCompraI = 0;

											if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicionArticulo = (int)dr.Cells["Posicion"].Value;
									
											oCmdActualiza.CommandText = string.Format("Select CSCostoCredito From Seriales Where idSerial = {0}", (int)dr.Cells["idSerial"].Value);
											decimal dPrecio = (decimal)oCmdActualiza.ExecuteScalar();
										
											oCmdActualiza.CommandText = string.Format("Exec GrabaDetalleIEBodega {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
												0, idCompraE, (int)dr.Cells["idArticulo"].Value, "", "", 1, dPrecio, dPrecio, iPosicionArticulo);
											idDetCompraE = (int)oCmdActualiza.ExecuteScalar();
										
											string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'MOVIMIENTO GENERADO AUTOMATICAMENTE DE CONSIGNACION', 3, 0, {4}, {5}", 
												(int)dr.Cells["idSerial"].Value, idDetCompraE, iBodegaCS, "", 0, (int)dr.Cells["Posicion"].Value);
											oCmdActualiza.CommandText = sSQLGDetalleSeriales;
											oCmdActualiza.ExecuteNonQuery();

											oCmdActualiza.CommandText = string.Format("Exec GrabaDetalleIEBodega {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
												0, idCompraI, (int)dr.Cells["idArticulo"].Value, "", "", 1, dPrecio, dPrecio, iPosicionArticulo);
											idDetCompraI = (int)oCmdActualiza.ExecuteScalar();
											
											oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'MOVIMIENTO GENERADO AUTOMATICAMENTE DE CONSIGNACION', 1, 0, {4}, {5}", 
												(int)dr.Cells["idSerial"].Value, idDetCompraI, (int)this.cmbBodega1.Value, "", 0, (int)dr.Cells["Posicion"].Value);;
											oCmdActualiza.ExecuteNonQuery();
										}
									}							
									#endregion Graba Registro

									#region Crea Asiento
									oCmdActualiza.CommandText = string.Format("Exec CreaAsiEgreso {0}", idCompraE);
									oCmdActualiza.ExecuteNonQuery();

									oCmdActualiza.CommandText = string.Format("Exec CreaAsiIngreso {0}", idCompraI);
									oCmdActualiza.ExecuteNonQuery();								
									#endregion Crea Asiento
								}
							}
						}
						#endregion Ingreso Egreso De Bodega Consignacion

						if (bNuevo && bNumeracionAutomatica)
						{
							string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);//(int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
						}
					
						#region Graba Factura
						#region Maestro
						int idDireccion = 0;
						int idCourier = 0;
						
																						
						if (this.cmbDirecciones.ActiveRow != null) idDireccion = (int)this.cmbDirecciones.Value;
						if (this.cmbCourier.ActiveRow != null) idCourier = (int)this.cmbCourier.Value;

						#region Fechas
						DateTime dtFecha = (DateTime)this.dtFecha.Value;
						DateTime dtFechaCad = DateTime.Today;
						DateTime dtFechaCad1 = DateTime.Today;
						DateTime dtFechaEnt = DateTime.Today;
						DateTime dtFechaEnt1 = DateTime.Today;
						DateTime dtHoraEntrega = DateTime.Today;
						DateTime dtHoraEntrega1 = DateTime.Today;
						bool bFecCad = false;
						bool bFecCad1 = false;
						bool bFecEnt = false;
						bool bHorEnt = false;
						bool bFecEnt1 = false;
						bool bHorEnt1 = false;
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
								dtFechaEnt = (DateTime) this.dtFechaEntrega.Value;
							}
							if (this.dtpHoraEntrega.Value != System.DBNull.Value) 
							{
								bHorEnt = true;
								dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
							}
							if (this.dtFechaEntregaCuriermbo1.Value != System.DBNull.Value) 
							{
								bFecEnt1 = true;
								dtFechaCad1 = (DateTime) this.dtFechaEntregaCuriermbo1.Value;
							}
							if (this.dtphoracurier.Value != System.DBNull.Value) 
							{
								bHorEnt1 = true;
								dtHoraEntrega1 = (DateTime) this.dtphoracurier.Value;
							}
						}
						if (bCrediPoint) bFecPrimPag = true;
						#endregion Fechas

						sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, {65}, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '', {54}, {55}, '{56}', '{57}', {58}, '{59}', {60}, {61}, '{62}', 0.00, {63}, '{64}','{66}','{67}'", 
							(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtIdSucursal.Value, //2
							(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, (int)this.txtNumEstado.Value, bConsignacion, //6
							(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, //10
							idVendedor, this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(), //13
							(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), //17
							dtFecha.ToString("yyyyMMdd"), this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), //21
							(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion, //25
							this.txtNotas.Text.ToString(), bCrediPoint, dTotalCuotas, NumCuotas, //29
							(decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"), //31
							(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, //34
							(decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, //37
							bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado, //42
							dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, AutorizaEntrada, dValorCuota, (int)this.txtIdEF.Value, //49
							(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, bVeriricador,//54
							(int)this.optTipoVenta.EditValue, this.txtNumTransaccion.Text.ToString(), this.txtDireccionEntrega.Text.ToString(), //57
							idCourier, this.txtNumGuia.Text.ToString(), Convert.ToDecimal(this.txtValorEntrega.Value), //60
							(int)this.txtNumIdproyectoZoom.Value, this.txtClienteNombre.Text.ToString(), iEncuesta, sOtrosEncuesta, (int)this.cmbBodega1.Value, 
							this.txtresponsablecurier.Text.ToString(),dtFechaCad1.ToString("yyyyMMdd")); 
						  
						oCmdActualiza.CommandText = sSQL;
						this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Maestro
						
						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							int iPosicion = 0;
							if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicion = (int)dr.Cells["Posicion"].Value;

							string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}",
								(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value,
								(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
								(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
								(bool)dr.Cells["EstadoSerial"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, (int)dr.Cells["Posicion"].Value, (bool)dr.Cells["Principal"].Value, 
								(bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value, (int)dr.Cells["idArticuloGE"].Value, (decimal)dr.Cells["Jornadas"].Value, (decimal)dr.Cells["Aporte"].Value, (int)dr.Cells["idSubProyecto"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	

						if (idTipoOrigen == 11 )
						{
							oCmdActualiza.CommandText = string.Format("Update Compra Set idOrigen = {1} Where idCompra = {0}", idProformaAseguradora, (int)this.txtNumIdCompra.Value);
							oCmdActualiza.ExecuteNonQuery();
						}
						
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);						
						this.ultraGrid1.DataSource = oDTDetalle;						
						#endregion Graba Factura
						
						#region Guarda la tabla de amortización de la factura
						if (bCrediPoint)
						{
							if (bNuevo && ((int) this.cmbFormaPago.Value == 9 || (int) this.cmbFormaPago.Value == 35))
							{
							
								decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
							
								DateTime dtFechaT = (DateTime)this.dtFecha.Value;
								string sSQLTA = string.Format("Exec GuardaTablaDeAmortizacion {0}, {1}, '{2}', '{3}', {4}, {5} , {6}", 
									dValor, NumCuotas, dtFechaT.ToString("yyyyMMdd"), dtFechPrimPago.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int) this.cmbFormaPago.Value, idCre_SolicitudWeb);						
								oCmdActualiza.CommandText = sSQLTA;
								oCmdActualiza.ExecuteNonQuery();

							}
						}
						#endregion Guarda la tabla de amortización de la factura

						#region Freelance
						
						//(bool) this.chkFreelance.Checked && 
						if (MenuLatinium.stIdDB == 1)
						{
							string sSQLFreelance = string.Format("Exec Com_FreelanceGeneraPeriodoCuotasTablaAmor {0}",
								(int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLFreelance;
							oCmdActualiza.ExecuteNonQuery();
						
							
							oCmdActualiza.CommandText = string.Format("Update Compra Set ifreelance = {1} Where idCompra = {0}",  (int)this.txtNumIdCompra.Value, Convert.ToInt32( this.chkFreelance.Checked));
							oCmdActualiza.ExecuteNonQuery();
						

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{
								if(Convert.ToInt32 (dr.Cells["idArticuloGE"].Value) > 0)
								{
									string sSQLDetalleFreelance = string.Format("Exec Com_FreelanceGeneraPeriodoCuotasTablaAmorGEX {0}, {1} ",
										(int)dr.Cells["idDetCompra"].Value, (int)this.txtNumIdCompra.Value);
									oCmdActualiza.CommandText = sSQLDetalleFreelance;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
						}

						
						#endregion Freelance

						if (bNuevo)
						{
							if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
							{
								oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", (int)this.txtNumIdCompra.Value);;
								string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

								string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
								string ValM11 = Modulo11(cadInvertida).ToString().Trim();
								string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

								oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.ExecuteNonQuery();
							}
						}

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
								dRetenido = dRetenido + Math.Round((decimal)dr.Cells["Valor"].Value, 2);
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
									string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", 
										(int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
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

								string sSQLConteoActivo = string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(Date, Fecha) <= '{1}' And Estado = 1", 
									(int)this.cmbBodega1.Value, dtFecha.ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLConteoActivo;
								int iConteoActivo = (int)oCmdActualiza.ExecuteScalar();

								if (iConteoActivo > 0)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
									{
										string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
											dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega1.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
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
															idSerial, (int)dr1.Cells["idArticulo"].Value, dr1.Cells["Serial"].Value.ToString(), (int)this.cmbBodega1.Value);
														oCmdActualiza.CommandText = sSQLGSeriales;
														int NIdSerial = (int)oCmdActualiza.ExecuteScalar();

														string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'LA SERIAL NO TIENE REGISTRO EN EL INVENTARIO.', 12, 0, {4}, {5}", 
															NIdSerial, IdDetCompra, (int)this.cmbBodega1.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int) dr1.Cells["Posicion"].Value);
														oCmdActualiza.CommandText = sSQLGDetalleSeriales;
														oCmdActualiza.ExecuteNonQuery();
													}
													else if (IEstado == 4 || IEstado == 2)
													{
														int idSerial = 0;
														int idOrigen = 0;

														if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
														if (dr1.Cells["idDetalleSerial"].Value != DBNull.Value) idOrigen = (int)dr1.Cells["idOrigen"].Value;
														
														string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", 
															idSerial, IdDetCompra, idOrigen, (int)this.cmbBodega1.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int) dr1.Cells["Posicion"].Value);
														oCmdActualiza.CommandText = sSQLGDetalleSeriales;
														oCmdActualiza.ExecuteNonQuery();

														#region RegistroCelulares
														oCmdActualiza.CommandText = string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",(int)dr1.Cells["idArticulo"].Value);
														int idArtGrup = (int)oCmdActualiza.ExecuteScalar();

														//int idArtGrup = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",(int)dr1.Cells["idArticulo"].Value));
											
														if ((int)this.cmbFormaPago.Value == 9  &&  idArtGrup==8 && (int)this.txtIdEF.Value == 14)
														{
															string sSQLRC = string.Format("Insert Into RegistroCelulares (idcompra, iddetcompra, idserial, serie) Values ({0}, {1}, {2},'{3}')", 
																(int) this.txtNumIdCompra.Value,(int)dr.Cells["idDetCompra"].Value,(int)dr1.Cells["idSerial"].Value,dr1.Cells["Serial"].Value.ToString());
															oCmdActualiza.CommandText = sSQLRC;
															oCmdActualiza.ExecuteNonQuery();
															PointMovil =1;
														}
														#endregion RegistroCelulares
															
														string sSQLActEstSerial = string.Format("Update Seriales Set Bodega = {2}, Estado = {0} Where idSerial = {1}", 
															4, idSerial, (int)this.cmbBodega1.Value);
														oCmdActualiza.CommandText = sSQLActEstSerial;															
														oCmdActualiza.ExecuteNonQuery();
																														
														string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
															dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega1.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
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

						#region Inventario
						oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 1", (int)this.txtNumIdCompra.Value);
						oCmdActualiza.ExecuteNonQuery();
						#endregion Inventario

						#region Estado Documento de origen
						if ((int)this.txtNumidOrigen.Value > 0)
						{
							if (idTipoOrigen == 43)
							{
								string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado = 5", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
								oCmdActualiza.CommandText = sSQLEstado;
								oCmdActualiza.ExecuteNonQuery();
							}
							else if (idTipoOrigen == 3 || idTipoOrigen == 11 || idTipoOrigen == 13)
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
							if(!bRebates)
							{
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
						}
						#endregion Pagos - Codigos de barras
												
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

						#region Garantia Extendida
						int iGarantiasExtendidas = 0;
						bool tienePointRespaldo = false;
						bool tieneAsistenciaSaludDental = false;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							string sNumeroGE = "";
							int idDetCompraGE = 0;
							int idSerialGE = 0;
							
							if ((int)dr.Cells["idArticulo"].Value == 12876)
							{
								tienePointRespaldo = true;
							}

							//if ((int)dr.Cells["idArticulo"].Value == 12920 || (int)dr.Cells["idArticulo"].Value == 12921)
							if ((int)dr.Cells["idArticulo"].Value == 12929 || (int)dr.Cells["idArticulo"].Value == 12930)	
							{
								tieneAsistenciaSaludDental = false;
							}

							if ((int)dr.Cells["idArticuloGE"].Value > 0)
							{								
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
								{
									if ((int)dr.Cells["idArticuloGE"].Value == (int)dr1.Cells["idArticulo"].Value)
									{
										idDetCompraGE = (int)dr1.Cells["idDetCompra"].Value;
										continue;
									}
								}

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.grdSerial.Rows.All)
								{
									if ((int)dr2.Cells["idDetCompra"].Value == idDetCompraGE)
									{
										idSerialGE = (int)dr2.Cells["idSerial"].Value;
										continue;
									}
								}

								if (bNuevo)
								{
									string sSQLNGE = string.Format("Exec NumeracionLocales 55, {0}, 1", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNGE;
									sNumeroGE = oCmdActualiza.ExecuteScalar().ToString();	
								}
								
								string sSQLGGE = string.Format("Exec GuardaGarantiaExtendida {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}", 
									0, (int)this.txtNumIdCompra.Value, (int)dr.Cells["idDetCompra"].Value, sNumeroGE, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)dr.Cells["idArticulo"].Value, 0, 
									(int)dr.Cells["idArticuloGE"].Value, idSerialGE, Convert.ToDecimal(dr.Cells["Precio"].Value), 1, idDetCompraGE);									
								oCmdActualiza.CommandText = sSQLGGE;
								oCmdActualiza.ExecuteNonQuery();

								iGarantiasExtendidas++;
							}
						}
						#endregion Garantia Extendida

						#region Bloqueos
						if (idBloqueaTransacciones > 0)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						if (bNuevo && (int)this.txtNumidOrigen.Value > 0)
						{
							string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
							oCmdActualiza.CommandText = sSQLBloqueoOrigen;
							oCmdActualiza.ExecuteNonQuery();
						}					
						#endregion Bloqueos

						#region Audita Descuento
						if (bDescuentoActivo)
						{
							string sMensaje = "";
					
							int iContador = 1;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
							{
								if (iContador > 1) sMensaje = " - ";

								sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

								iContador++;
							}

							string sSQLCD = string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'FACTURA', {1}, 3, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje);
							oCmdActualiza.CommandText = sSQLCD;
							oCmdActualiza.ExecuteNonQuery();						
						}
						#endregion Audita Descuento

						string sSQLCognoScoring = string.Format("Exec Cre_ActualizaEstadoDatosGenerales '{0}', {1}", this.txtRuc.Text.ToString(),(int)this.txtNumIdCompra.Value );
						oCmdActualiza.CommandText = sSQLCognoScoring;
						oCmdActualiza.ExecuteNonQuery();

						if (bRebates && idLiquidacionRebates > 0)
						{
							string sSQLUpd = string.Format( "UPDATE Rebates_Liquidacion SET Estado = 2, idCompra = {0} WHERE idRebates_Liquidacion = {1}",(int)this.txtNumIdCompra.Value,idLiquidacionRebates);
							oCmdActualiza.CommandText = sSQLUpd;
							oCmdActualiza.ExecuteNonQuery();
						}

						oTransaction.Commit();

						oConexion.Close();

						#region FamilyFriends Credito
						FamilyFriendsCompraVenta(FamilyFriendsService.EstadoFFCompraContabilizada);
						#endregion FamilyFriends Credito


						Cursor = Cursors.Default;							

						Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, (int)this.txtNumIdCompra.Value, cdsSeteoF);


						#region Impresión
						if(!bRebates) ImpresionDirecta(iGarantiasExtendidas);

						if(tienePointRespaldo)
						{
							ImpresionPointRespaldo();
						}
						if(tieneAsistenciaSaludDental)
						{
							ImpresionAsistenciaSaludDental();
						}
						if ((int)this.cmbFormaPago.Value == 9) 
						{
							Facturacion.ImpresionRecibos(iAnticipoCDP, 12, cdsSeteoF);

							#region Imprime Celulares 
							if ((int)this.txtIdEF.Value ==14)
							{
								ParameterFields paramFields = new ParameterFields ();
				
								ParameterField paramFieldId = new ParameterField ();
								ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
								paramFieldId.ParameterFieldName = "@idCompra";
								ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
								paramFieldId.CurrentValues.Add (ValIdId);
								paramFields.Add (paramFieldId);

								Reporte miRep = new Reporte("ContratoCompraVentaCelular.rpt", "");
								miRep.MdiParent = this.MdiParent;
								miRep.Titulo("Recepción Mercaderia de Retiro");
								miRep.crVista.ParameterFieldInfo = paramFields;
								miRep.Show();	
							}
							#endregion Imprime Celulares 
						}
						#endregion Impresión
							
						#region Controles
						if ((int)this.txtNumEstado.Value == 11)	this.lblEstado.Text = "MANUAL";
						else if ((int)this.txtNumEstado.Value == 10) this.lblEstado.Text = "CONTABILIZADO";

						this.cmbFormaPago.Enabled = false;
						this.cmbBodega.Enabled = false;
						this.cmbBodega1.Enabled = false;
						this.cmbVendedor.Enabled = false;
						this.cmbTipoRuc.Enabled = false;

						this.cmbFechaPrimerPago.Enabled = false;
						this.cmbNumContadoCredito.Enabled = false;

						this.txtNumero.Enabled = false;
						this.txtRuc.Enabled = false;
						this.txtRuc1.Enabled = false;
						this.dtFecha.Enabled = false;

						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
						if (this.chkEntregado.Checked) this.btnImprimeED.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (!miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
						if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						this.btnGuardar.Enabled = false;
						this.btnCancelar.Enabled = true;						
						this.btnClaveDescuento.Enabled = false;
						this.btnTemporal.Enabled = false;
						this.btnCliente.Enabled = false;
						this.btnCliente2.Enabled = false;
						this.btnValidarMail.Enabled = false;
						this.txtClienteNombre.Enabled = false;

						
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.AllowEdit;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].CellActivation = Activation.AllowEdit;

						HabilitaTab(false);
						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
						FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);
						
						this.grdSerial.Enabled = true;
						bNuevo = false;
						bEdicion = false;
						bDescuentoActivo = false;
						idBloqueaTransacciones = 0;
						dClavePorcDscto = 0.00m;
						idCuponWeb = 0;
						bCuponWeb = false;
						bValidaTipoVenta = false;						
						#endregion Controles					

						if (idCompraOrigen > 0) this.DialogResult = DialogResult.OK;

						#endregion proceso 
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
				if (bServicioTecnico) this.DialogResult = DialogResult.Yes;
				return;
				//					}
				//				}
			}
			#endregion Credito - CrediPOINT SIN ENTRADA - PROMOCIONALES

			bool bDYD = false;

			#region Efectivo - Tarjeta
			this.txtNumEstado.Value = 10;
			
			bool bEmpleado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaEmpleado '{0}'", this.txtRuc.Text.ToString()), true);
			int idPersonal = 0;
			if (bEmpleado)
				idPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionConsultaEmpleado '{0}'", this.txtRuc.Text.ToString()), true);
			
			if ((int)this.cmbFormaPago.Value == 21)
			{
				this.txtCuotaDeEntrada.Value = 0.00m;
			}			
			if (bCrediPoint) dValorCrediPoint = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
			
			using (VentaFormaDePago miCobro = new VentaFormaDePago ((int)this.txtNumIdCompra.Value, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, this.txtComprobante.Text.ToString(), BUsoNotaDeCredito, idCompraOrigen, dSaldoNotaDeCredito, bRetencionCaja, dValorRetencion, bUsoAnticipo, dAnticipo, bCrediPoint, dValorCrediPoint, (decimal)this.txtCuotaDeEntrada.Value, 1, (int)this.txtNumIdCliente.Value, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, (int)this.cmbBodega.Value, this.ultraGrid1, (DateTime)this.dtFecha.Value, bDYD, idPersonal, DMonto))
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

							#region Ingreso Egreso De Bodega Consignacion
							if (MenuLatinium.bConsignacionVenta)
							{
								if (bNuevo)
								{
									int iCont = 0;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
										if ((bool)dr.Cells["Consignacion"].Value) iCont++;

									if (iCont > 0)
									{
										int idCompraE = 0; 
										int idCompraI = 0;
										string sNumeroIB = ""; 
										string sNumeroEB = "";
										int iBodegaCS = 0;

										oCmdActualiza.CommandText = string.Format("Select Bodega From AsignaBodegasLocales Where Local = {0} And Consignacion = 1", (int)this.cmbBodega.Value);
										iBodegaCS = (int)oCmdActualiza.ExecuteScalar();
								
										#region Numeracion
										if (bNuevo)
										{									
											oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 8, {0}, 0", iBodegaCS);
											sNumeroEB = oCmdActualiza.ExecuteScalar().ToString();

											oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 9, {0}, 0", (int)this.cmbBodega1.Value);
											sNumeroIB = oCmdActualiza.ExecuteScalar().ToString();
										}
										#endregion Numeracion

										#region Graba Registro 
										oCmdActualiza.CommandText = string.Format("Exec GrabaIEBodega {0}, 8, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, '{8}', '', 7",
											idCompraE, 171570, iBodegaCS, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumeroEB, 3, 0.00m, 0.00m, this.txtNotas.Text.ToString());;
										idCompraE = (int)oCmdActualiza.ExecuteScalar();

										oCmdActualiza.CommandText = string.Format("Exec GrabaIEBodega {0}, 9, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, '{8}', '', 7",
											idCompraI, 171570, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumeroIB, 3, 0.00m, 0.00m, this.txtNotas.Text.ToString());;
										idCompraI = (int)oCmdActualiza.ExecuteScalar();
													
										int iPosicionArticulo = 0;
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdSerial.Rows.All)
										{
											if ((bool)dr.Cells["Consignacion"].Value)
											{
												int idDetCompraE = 0;
												int idDetCompraI = 0;

												if (dr.Cells["Posicion"].Value != System.DBNull.Value) iPosicionArticulo = (int)dr.Cells["Posicion"].Value;
									
												oCmdActualiza.CommandText = string.Format("Select CSCostoCredito From Seriales Where idSerial = {0}", (int)dr.Cells["idSerial"].Value);
												decimal dPrecio = (decimal)oCmdActualiza.ExecuteScalar();
										
												oCmdActualiza.CommandText = string.Format("Exec GrabaDetalleIEBodega {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
													0, idCompraE, (int)dr.Cells["idArticulo"].Value, "", "", 1, dPrecio, dPrecio, iPosicionArticulo);
												idDetCompraE = (int)oCmdActualiza.ExecuteScalar();
																					
												oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'MOVIMIENTO GENERADO AUTOMATICAMENTE DE CONSIGNACION', 3, 0, {4}, {5}", 
													(int)dr.Cells["idSerial"].Value, idDetCompraE, iBodegaCS, "", 0, (int)dr.Cells["Posicion"].Value);;
												oCmdActualiza.ExecuteNonQuery();

												oCmdActualiza.CommandText = string.Format("Exec GrabaDetalleIEBodega {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
													0, idCompraI, (int)dr.Cells["idArticulo"].Value, "", "", 1, dPrecio, dPrecio, iPosicionArticulo);
												idDetCompraI = (int)oCmdActualiza.ExecuteScalar();
											
												oCmdActualiza.CommandText = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'MOVIMIENTO GENERADO AUTOMATICAMENTE DE CONSIGNACION', 1, 0, {4}, {5}", 
													(int)dr.Cells["idSerial"].Value, idDetCompraI, (int)this.cmbBodega1.Value, "", 0, (int)dr.Cells["Posicion"].Value);
												oCmdActualiza.ExecuteNonQuery();
											}
										}						
										#endregion Graba Registro

										#region Crea Asiento
										oCmdActualiza.CommandText = string.Format("Exec CreaAsiEgreso {0}", idCompraE);
										oCmdActualiza.ExecuteNonQuery();
									
										oCmdActualiza.CommandText = string.Format("Exec CreaAsiIngreso {0}", idCompraI);
										oCmdActualiza.ExecuteNonQuery();								
										#endregion Crea Asiento
									}
								}
							}
							#endregion Ingreso Egreso De Bodega Consignacion

							if (bNuevo && bNumeracionAutomatica)
							{
								string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 1", (int)this.cmbBodega.Value);//(int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLNumero;
								this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
							}	
			
							#region Graba Factura
							#region Maestro
							int idDireccion = 0;
							int idCourier = 0;
							if (this.cmbDirecciones.ActiveRow != null) idDireccion = (int)this.cmbDirecciones.Value;
							if (this.cmbCourier.ActiveRow != null) idCourier = (int)this.cmbCourier.Value;

							#region Fechas
							DateTime dtFecha = (DateTime)this.dtFecha.Value;
							DateTime dtFechaCad = DateTime.Today;
							DateTime dtFechaEnt = DateTime.Today;
							DateTime dtHoraEntrega = DateTime.Today;
							bool bFecCad = false;
							bool bFecEnt = false;
							bool bHorEnt = false;
							bool bFecPrimPag = false;
							int iTipoVenta = (int)this.optTipoVenta.EditValue;
							if (this.cmbCaducidad.Value != System.DBNull.Value) 
							{
								bFecCad = true;
								dtFechaCad = (DateTime) this.cmbCaducidad.Value;
							}
							if (this.dtFechaEntrega.Value != System.DBNull.Value) 
							{
								bFecEnt = true;
								dtFechaEnt = (DateTime) this.dtFechaEntrega.Value;
							}
							if (this.dtpHoraEntrega.Value != System.DBNull.Value) 
							{
								bHorEnt = true;
								dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
							}
							if (bCrediPoint) bFecPrimPag = true;
							DateTime dtFechaCourier = DateTime.Today;
							if (this.dtFechaEntregaCuriermbo1.Value != System.DBNull.Value) dtFechaCourier = (DateTime)this.dtFechaEntregaCuriermbo1.Value;
							#endregion Fechas

							sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, {67}, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '', {54}, {55}, '{56}', '{57}', {58}, '{59}', {60}, {61}, '{62}', 0.00, {63}, '{64}', '{65}', '{66}'", 
								(int) this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value,
								(int)this.txtIdSucursal.Value, 
								(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
								(int)this.txtNumEstado.Value,
								bConsignacion,
								(int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
								(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
								idVendedor, this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
								(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
								dtFecha.ToString("yyyyMMdd"), 
								this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad.ToString("yyyyMMdd"), //21
								(bool)this.chkEntregado.Checked, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), idDireccion,
								this.txtNotas.Text.ToString(),
								bCrediPoint, dTotalCuotas, NumCuotas, (decimal)this.txtCuotaDeEntrada.Value, dtFechPrimPago.ToString("yyyyMMdd"),
								(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
								bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado, //42
								dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, AutorizaEntrada, dValorCuota, (int)this.txtIdEF.Value,
								(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, bVeriricador,
								iTipoVenta, this.txtNumTransaccion.Text.ToString(), this.txtDireccionEntrega.Text.ToString(), idCourier, this.txtNumGuia.Text.ToString(), Convert.ToDecimal(this.txtValorEntrega.Value),
								(int)this.txtNumIdproyectoZoom.Value, this.txtClienteNombre.Text.ToString(),
								iEncuesta, sOtrosEncuesta, this.txtresponsablecurier.Text.ToString(), dtFechaCourier.ToString("yyyyMMdd HH:mm"), (int)this.cmbBodega1.Value);
							//@Responsablecurier varchar(100)='', @FecEntCurier Datetime = '20000101', @HorEntCurier Datetime = '20000101'
							oCmdActualiza.CommandText = sSQL;
							this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
							#endregion Maestro

							#region Detalle
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{	
								string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}",
									(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value,
									(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
									(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
									(bool)dr.Cells["EstadoSerial"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, (int)dr.Cells["Posicion"].Value, (bool)dr.Cells["Principal"].Value, 
									(bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value, (int)dr.Cells["idArticuloGE"].Value, (decimal)dr.Cells["Jornadas"].Value, (decimal)dr.Cells["Aporte"].Value, (int)dr.Cells["idSubProyecto"].Value);
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

							#region Graba Google
							if ((int)this.cmbFormaPago.Value == 32)
							{
								if (sUrlNube.Length > 10)
								{
									string sSQLGoogle = string.Format("Exec GrabaUrlNubeDeuna {0}, '{1}', '{2}'", (int)this.txtNumIdCompra.Value, "", sUrlNube);
									oCmdActualiza.CommandText = sSQLGoogle;
									oCmdActualiza.ExecuteNonQuery();
								}
							}											
							#endregion Graba Google

							#region Guarda la tabla de amortización de la factura
							if (bCrediPoint)
							{
								if (bNuevo && (int) this.cmbFormaPago.Value == 9)
								{
									decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
							
									DateTime dtFechaT = (DateTime)this.dtFecha.Value;
									string sSQLTA = string.Format("Exec GuardaTablaDeAmortizacion {0}, {1}, '{2}', '{3}', {4}", 
										dValor, NumCuotas, dtFechaT.ToString("yyyyMMdd"), dtFechPrimPago.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value);						
									oCmdActualiza.CommandText = sSQLTA;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							#endregion Guarda la tabla de amortización de la factura
							
							#region Clave S.R.I.
							if (bNuevo)
							{
								if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
								{
									oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", (int)this.txtNumIdCompra.Value);;
									string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

									string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
									string ValM11 = Modulo11(cadInvertida).ToString().Trim();
									string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

									oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, (int)this.txtNumIdCompra.Value);
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							#endregion Clave S.R.I.

							#region Freelance
							//(bool) this.chkFreelance.Checked && 
							if (MenuLatinium.stIdDB == 1)
							{
								string sSQLFreelance = string.Format("Exec Com_FreelanceGeneraPeriodoCuotasTablaAmor {0}",
									(int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLFreelance;
								oCmdActualiza.ExecuteNonQuery();
						
							
								oCmdActualiza.CommandText = string.Format("Update Compra Set ifreelance = {1} Where idCompra = {0}",  (int)this.txtNumIdCompra.Value, Convert.ToInt32( this.chkFreelance.Checked));
								oCmdActualiza.ExecuteNonQuery();
						

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
								{
									if(Convert.ToInt32 (dr.Cells["idArticuloGE"].Value) > 0)
									{
										string sSQLDetalleFreelance = string.Format("Exec Com_FreelanceGeneraPeriodoCuotasTablaAmorGEX {0}, {1} ",
											(int)dr.Cells["idDetCompra"].Value, (int)this.txtNumIdCompra.Value);
										oCmdActualiza.CommandText = sSQLDetalleFreelance;
										oCmdActualiza.ExecuteNonQuery();
									}
								}
							}						
							#endregion Freelance

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
									dRetenido = dRetenido + Math.Round(Convert.ToDecimal(dr.Cells["Valor"].Value), 2);
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
								
									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}, {19}", // {19}", 
										idDetFormaPago, 
										(int)this.txtNumIdCompra.Value, 1, 
										(int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
										idBanco, idTarjeta,
										dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),								
										dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, 
										dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
										idPlan, idPlazo, dComTarjeta, idDocumento, (bool)dr.Cells["Temporal"].Value);
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

							#region Inventario
							oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 1", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.ExecuteNonQuery();
							#endregion Inventario

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
										string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", 
											(int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
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

									string sSQLConteoActivo = string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(Date, Fecha) <= '{1}' And Estado = 1", 
										(int)this.cmbBodega1.Value, dtFecha.ToString("yyyyMMdd"));
									oCmdActualiza.CommandText = sSQLConteoActivo;
									int iConteoActivo = (int)oCmdActualiza.ExecuteScalar();

									if (iConteoActivo > 0)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in grdSerial.Rows.All)
										{
											string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
												dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega1.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
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
																string sSQLGSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', 12, {3}", 
																	idSerial, (int)dr1.Cells["idArticulo"].Value, dr1.Cells["Serial"].Value.ToString(), (int)this.cmbBodega1.Value);
																oCmdActualiza.CommandText = sSQLGSeriales;
																NIdSerial = (int)oCmdActualiza.ExecuteScalar();
															}
															else NIdSerial = idSerial;

															string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, 0, {2}, '{3}', 1, 'LA SERIAL NO TIENE REGISTRO EN EL INVENTARIO.', 12, 0, {4}, {5}", 
																NIdSerial, IdDetCompra, (int)this.cmbBodega1.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
															oCmdActualiza.CommandText = sSQLGDetalleSeriales;
															oCmdActualiza.ExecuteNonQuery();

															oCmdActualiza.CommandText = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", 12, idSerial, (int)this.cmbBodega1.Value);
															oCmdActualiza.ExecuteNonQuery();

															string sSQLSM = string.Format("Insert Into SerialesManuales (idSerial, Serial, idArticulo, Bodega) Values ({0}, '{1}', {2}, {3})", 
																NIdSerial, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value, (int)this.cmbBodega1.Value);
															oCmdActualiza.CommandText = sSQLSM;
															oCmdActualiza.ExecuteNonQuery();	
														}
														else if (IEstado == 4 || IEstado == 2)
														{
															int idSerial = 0;
															int idOrigen = 0;

															if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
															if (dr1.Cells["idOrigen"].Value != DBNull.Value) idOrigen = (int)dr1.Cells["idOrigen"].Value;

															string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", 
																idSerial, IdDetCompra, idOrigen, (int)this.cmbBodega1.Value, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
															oCmdActualiza.CommandText = sSQLGDetalleSeriales;
															oCmdActualiza.ExecuteNonQuery();

															#region RegistroCelulares
															oCmdActualiza.CommandText = string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",(int)dr1.Cells["idArticulo"].Value);
															int idArtGrup = (int)oCmdActualiza.ExecuteScalar();

															//int idArtGrup = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",(int)dr1.Cells["idArticulo"].Value));
											
															if ((int)this.cmbFormaPago.Value == 9  &&  idArtGrup==8 && (int)this.txtIdEF.Value == 14 )
															{
																string sSQLRC = string.Format("Insert Into RegistroCelulares (idcompra, iddetcompra, idserial, serie) Values ({0}, {1}, {2},'{3}')", 
																	(int) this.txtNumIdCompra.Value,(int)dr.Cells["idDetCompra"].Value,(int)dr1.Cells["idSerial"].Value,dr1.Cells["Serial"].Value.ToString());
																oCmdActualiza.CommandText = sSQLRC;
																oCmdActualiza.ExecuteNonQuery();
																PointMovil =1;
															}
															#endregion RegistroCelulares
																
															string sSQLActEstSerial = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", 
																4, idSerial, (int)this.cmbBodega1.Value);
															oCmdActualiza.CommandText = sSQLActEstSerial;																
															oCmdActualiza.ExecuteNonQuery();

															string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
																dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega1.Value, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
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
								else if (idTipoOrigen == 3 || idTipoOrigen == 11 || idTipoOrigen == 13)
								{
									string sSQLEstado = string.Format("Update Compra Set Estado = 9, Comprobante = '{2}' Where idCompra = {0} And idTipoFactura = {1} And Estado In (3, 5)", (int)this.txtNumidOrigen.Value, idTipoOrigen, this.txtNumero.Text.ToString());
									oCmdActualiza.CommandText = sSQLEstado;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							#endregion Estado Documento de origen
							
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

							#region Garantia Extendida
							int iGarantiasExtendidas = 0;
							bool tieneAsistenciaSaludDental = false;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{
								string sNumeroGE = "";
								int idDetCompraGE = 0;
								int idSerialGE = 0;
	
								//if ((int)dr.Cells["idArticulo"].Value == 12920 || (int)dr.Cells["idArticulo"].Value == 12921)
								if ((int)dr.Cells["idArticulo"].Value == 12929 || (int)dr.Cells["idArticulo"].Value == 12930)
								{
									tieneAsistenciaSaludDental = false;
								}

								if ((int)dr.Cells["idArticuloGE"].Value > 0)
								{								
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
									{
										if ((int)dr.Cells["idArticuloGE"].Value == (int)dr1.Cells["idArticulo"].Value)
										{
											idDetCompraGE = (int)dr1.Cells["idDetCompra"].Value;
											continue;
										}
									}

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.grdSerial.Rows.All)
									{
										if ((int)dr2.Cells["idDetCompra"].Value == idDetCompraGE)
										{
											idSerialGE = (int)dr2.Cells["idSerial"].Value;
											continue;
										}
									}

									if (bNuevo)
									{
										string sSQLNGE = string.Format("Exec NumeracionLocales 55, {0}, 1", (int)this.cmbBodega.Value);
										oCmdActualiza.CommandText = sSQLNGE;
										sNumeroGE = oCmdActualiza.ExecuteScalar().ToString();	
									}
								
									string sSQLGGE = string.Format("Exec GuardaGarantiaExtendida {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}", 
										0, (int)this.txtNumIdCompra.Value, (int)dr.Cells["idDetCompra"].Value, sNumeroGE, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)dr.Cells["idArticulo"].Value, 0, 
										(int)dr.Cells["idArticuloGE"].Value, idSerialGE, Convert.ToDecimal(dr.Cells["Precio"].Value), 1, idDetCompraGE);									
									oCmdActualiza.CommandText = sSQLGGE;
									oCmdActualiza.ExecuteNonQuery();

									iGarantiasExtendidas++;
								}
							}
							#endregion Garantia Extendida

							#region Bloqueos
							if (idBloqueaTransacciones > 0)
							{
								string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQLBloqueo;
								oCmdActualiza.ExecuteNonQuery();
							}
							if (bNuevo && (int)this.txtNumidOrigen.Value > 0)
							{
								string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
								oCmdActualiza.CommandText = sSQLBloqueoOrigen;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Bloqueos

							#region Audita Descuento
							if (bDescuentoActivo)
							{
								string sMensaje = "";
					
								int iContador = 1;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
								{
									if (iContador > 1) sMensaje = " - ";

									sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

									iContador++;
								}

								string sSQLCD = string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'FACTURA', {1}, 3, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje);
								oCmdActualiza.CommandText = sSQLCD;
								oCmdActualiza.ExecuteNonQuery();						
							}
							#endregion Audita Descuento
							
							if (this.chkEntregado.Checked)
							{
								string sSQLCE = string.Format("Exec GuardaSeguimientoWEB {0}, {1}, '{2}', '{3}', {4}", 
									0, (int)this.txtNumIdCompra.Value, Convert.ToDateTime(this.dtFechaEntregaCuriermbo1.Value).ToString("yyyyMMdd HH:mm"), this.txtDireccionEntrega.Text.ToString(), 1);
								oCmdActualiza.CommandText = sSQLCE;
								oCmdActualiza.ExecuteNonQuery();	
							}

							if (idMail > 0)
							{
								string sSQLLM = string.Format("Update listacorreosmarketing Set idCompra = {0} Where idCorreosMarketing = {1}", (int)this.txtNumIdCompra.Value, idMail);
								oCmdActualiza.CommandText = sSQLLM;
								oCmdActualiza.ExecuteNonQuery();
							}
							if (BPagados == true)
							{
								string sSQP2 = string.Format("Exec ActualizaDatosPagados {0}, {1}", 
									IdPagados, (int)this.txtNumIdCompra.Value);
								oCmdActualiza.CommandText = sSQP2;
								oCmdActualiza.ExecuteNonQuery();	
							}

							#region Nota De Credito
							if (idCompraOrigen > 0)
							{
								oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 1, 1";
								string sNumeroNC = oCmdActualiza.ExecuteScalar().ToString();

								oCmdActualiza.CommandText = string.Format("Exec CreaNotaDeCreditoCambioProducto {0}, '{1}', {2}, '{3}'", 
									idCompraOrigen, sNumeroNC, idMotivo, "");
								int idNotaDeCredito = (int)oCmdActualiza.ExecuteScalar();

								//								oCmdActualiza.CommandText = string.Format("Exec CreaAsientoNotasDeCredito {0}", idNotaDeCredito);
								//								oCmdActualiza.ExecuteNonQuery();

								oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 5", idNotaDeCredito);;
								string sClaveAccesoNC = oCmdActualiza.ExecuteScalar().ToString();

								string cadInvertidaNC = Facturacion.invertirCadena(sClaveAccesoNC.Substring(0, 48));
								string ValM11NC = Facturacion.Modulo11(cadInvertidaNC).ToString().Trim();
								string sClaveSRINC = sClaveAccesoNC.ToString().Substring(0, 48) + ValM11NC;

								oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRINC, idNotaDeCredito);
								oCmdActualiza.ExecuteNonQuery();

								if (dTotalOrigen > Convert.ToDecimal(this.txtTotal.Value))
								{
									decimal dDiferencia = dTotalOrigen - Convert.ToDecimal(this.txtTotal.Value);

									string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 14, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'DEVOLUCION DE DINERO CON NOTA DE CREDITO', 0", 
										idNotaDeCredito, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Math.Round(dDiferencia, 2));
									oCmdActualiza.CommandText = sSQLAP;
									oCmdActualiza.ExecuteNonQuery();
								}

								MessageBox.Show(string.Format("Se genero con exito la nota de credito N.: '{0}'", sNumeroNC));
							}
							#endregion Nota De Credito

							oTransaction.Commit();

							oConexion.Close();

							#region FamilyFriends Efectivo Tarjeta
							FamilyFriendsCompraVenta(FamilyFriendsService.EstadoFFCompraContabilizada);
							#endregion FamilyFriends Efectivo Tarjeta
					
							#region Impresión
							ImpresionDirecta(0);//iGarantiasExtendidas

							if(tieneAsistenciaSaludDental)
							{
								ImpresionAsistenciaSaludDental();
							}
						
							if ((int)this.cmbFormaPago.Value == 9) 
							{
								Facturacion.ImpresionRecibos( iAnticipoCDP, 12, cdsSeteoF);
								#region Imprime Celulares 
								if ((int)this.txtIdEF.Value ==14)
								{
									ParameterFields paramFields = new ParameterFields ();
				
									ParameterField paramFieldId = new ParameterField ();
									ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
									paramFieldId.ParameterFieldName = "@idCompra";
									ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
									paramFieldId.CurrentValues.Add (ValIdId);
									paramFields.Add (paramFieldId);

									Reporte miRep = new Reporte("ContratoCompraVentaCelular.rpt", "");
									miRep.MdiParent = this.MdiParent;
									miRep.Titulo("Recepción Mercaderia de Retiro");
									miRep.crVista.ParameterFieldInfo = paramFields;
									miRep.Show();	
								}
								#endregion Imprime Celulares 
							}
							#endregion Impresión

							Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, (int)this.txtNumIdCompra.Value, cdsSeteoF);
							
							if ((int)this.txtNumEstado.Value == 11)	this.lblEstado.Text = "MANUAL";
							else if ((int)this.txtNumEstado.Value == 10) this.lblEstado.Text = "CONTABILIZADO";

							#region Bono entrada actualizar fecha uso bono
							int idBonoEntradaTipo = 0;
							int idBonoEntradaIdPrefactura = 0;
							string squery = String.Format("Exec BonoEntradaConsultarIdPrefactura {0}", (int)this.txtNumIdCompra.Value);
							idBonoEntradaIdPrefactura = Funcion.iEscalarSQL(cdsSeteoF, squery);

							//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
							DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
							if (dvResultado != null)
							{
								foreach (DataRowView rowView in dvResultado)
								{
									DataRow row = rowView.Row;
									idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
									//Fecha Uso Bono Entrada 
									//Validar si el idOrigen, pertenecese a la tabla BonoEntrada, con Bono 25
									if(idBonoEntradaIdPrefactura > 0)
									{
										//Actualizar fecha bono uso
										squery = String.Format("Exec BonoEntradaFechaUsoBonoAct {0},{1}", idBonoEntradaIdPrefactura, idBonoEntradaTipo);
										Funcion.EjecutaSQL(cdsSeteoF,squery);
									}
								}
							}
							#endregion Bono entrada actualizar fecha uso bono
					
							#region Controles
							this.cmbFormaPago.Enabled = false;
							this.cmbBodega.Enabled = false;
							this.cmbBodega1.Enabled = false;
							this.cmbVendedor.Enabled = false;
							this.cmbTipoRuc.Enabled = false;

							this.txtNumero.Enabled = false;
							this.txtRuc.Enabled = false;
							this.txtRuc1.Enabled = false;
							this.dtFecha.Enabled = false;

							if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
							if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
							if (miAcceso.Editar) this.btnEditar.Enabled = true;
							if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
							else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)							
								if ((int)this.cmbFormaPago.Value != 9)								
									if ((DateTime)this.dtFecha.Value == DateTime.Today) this.btnAnular.Enabled = true;
								
							if (!miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							this.btnGuardar.Enabled = false;
							this.btnCancelar.Enabled = true;							
							this.btnClaveDescuento.Enabled = false;
							this.btnTemporal.Enabled = false;
							this.btnCliente.Enabled = false;
							this.btnCliente2.Enabled = false;
							this.btnValidarMail.Enabled = false;
							this.txtClienteNombre.Enabled = false;

							HabilitaTab(false);

							FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
							FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);							
							
							bNuevo = false;
							bEdicion = false;
							bDescuentoActivo = false;
							idBloqueaTransacciones = 0;
							dClavePorcDscto = 0.00m;
							idCuponWeb = 0;
							bCuponWeb = false;
							bValidaTipoVenta = false;

							if (bServicioTecnico) this.DialogResult = DialogResult.Yes;
							#endregion Controles

							if (idCompraOrigen > 0) this.DialogResult = DialogResult.OK;
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
				else
				{		
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaAnticipos {0}, {1}, {2}, '', 0", iAnticipoCDP, iAsientoCDP, 1), true);
				
					if (bDYD)
					{
						MessageBox.Show("Esta factura contiene Proteccion D & D, debe volver a generar el documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
						this.btnCancelar_Click(sender, e);						
					}
					else bEdicion = true;
				}
			}
			#endregion Efectivo - Tarjeta
			#endregion Grabar
		}

		private void FamilyFriendsCompraVenta(int Estado)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "FamilyFriendsCompra_Venta";

			int idCompra = (int)this.txtNumIdCompra.Value;
			bool debeGuardar = true;

			bool estadoNeutro = (!estadoFF.EsFF && !estadoFF.RechazoPromocion);

			// =====================================================
			// 0) VALIDACIÓN PREVIA (ESTADO DE PANTALLA)
			// =====================================================
			if (estadoNeutro)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Estado neutro FF en pantalla, se reconsulta BD",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				// Re-consulta a BD para confirmar estado real
				ConsultarEstadoFamilyFriends(idCompra);

				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					string.Format(
					"Post reconsulta FF | EsFF: {0}, RechazoPromocion: {1}, IdWEB_FF_Invitado: {2}",
					estadoFF.EsFF,
					estadoFF.RechazoPromocion,
					estadoFF.IdWEB_FF_Invitado
					),
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				// recalcular estado neutro tras reconsulta
				estadoNeutro = (!estadoFF.EsFF && !estadoFF.RechazoPromocion);
			}

			// =====================================================
			// 1) VALIDACIONES DEFINITIVAS
			// =====================================================
			if (estadoNeutro)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Estado neutro FF confirmado, no se guarda",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				debeGuardar = false;
			}

			// =====================================================
			// 2) DECISIÓN FINAL
			// =====================================================
			if (!debeGuardar)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Fin proceso FF sin guardar",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);
				return;
			}

			// =====================================================
			// 3) PREPARAR MODELO
			// =====================================================
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			FamilyFriendsCompraModel model = new FamilyFriendsCompraModel();
			model.IdCompra = idCompra;
			model.IdWEB_FF_Invitado = estadoFF.IdWEB_FF_Invitado;
			model.IdTipoPrecio = idTipoPrecio;
			model.IdFormaPago = (int)this.cmbFormaPago.Value;
			model.IdCliente = (int)this.txtNumIdCliente.Value;

			model.Estado = estadoFF.RechazoPromocion
				? FamilyFriendsService.EstadoFFCompraNoAcepta
				: Estado;

			// =====================================================
			// 4) GUARDAR
			// =====================================================
			Guid id = service.GuardarCompra(model);

			Funcion.LogSistema(frmRequerimiento.sconexionInicio,
				nombreClase + ": " + nombreMetodo,
				"GuardarCompra_Venta OK | idWEB_FF_Compra: " + id,
				nombreMetodo,
				Funcion.slogtipoinformacion
				);
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{				
			this.Consultar((int)this.txtNumIdCompra.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede editar FACTURA ANULADA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;		
			}

			if (!miAcceso.Editar)
			{								
				MessageBox.Show("No tiene Acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				return;
			}

			if ((int)this.txtNumEstado.Value == 6)
			{
				if ((int)this.cmbBodega.Value != 1 && (int)this.cmbBodega.Value != 101 && (int)this.cmbBodega.Value != 153)
				{
					int iTransferencias = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CompraTransf Where Generado = 1 And Bodega = {0}", (int)this.cmbBodega.Value));

					if (iTransferencias > 0)
					{
						MessageBox.Show(string.Format("No puede crear facturas hasta que procese la(s) {0} transferencias pendientes", iTransferencias), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.btnCancelar_Click(sender, e);
						return;
					}
				}
			}

			if ((int)this.txtNumEstado.Value != 6 && !miAcceso.EditarContabilizado)
			{								
				MessageBox.Show("No puede editar una Factura Contabilizada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}			

			if (!miAcceso.EditaImpresa && bImpreso)
			{
				MessageBox.Show("No puede Editar Facturas Impresas.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFE_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) == 2)
			{
				MessageBox.Show("No puede Editar la Factura esta autorizada por el S.R.I.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("No puede Editar la Factura esta Asociada a una Retencion que a Afecta a Multiples Facturas, \n\nAnule la Retencion o Consulte al Administrador", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Bloquear, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true))
			{
				MessageBox.Show("No puede Editar la Factura esta Bloqueada", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta")) 
			{				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Editar esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);												
				return;
			}
      			
			if (!FuncionesInventario.ValidaBodegaInventarios(this.ultraGrid1, (int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;

			DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
			
			if ((miAcceso.EditarSoloDia && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Estado, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) != 6) && dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
			{
				MessageBox.Show("Solo puede editar Facturas de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}

			string sNumeroNC = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VentasValidaNotaDeCredito {0}", (int)this.txtNumIdCompra.Value));
			if (sNumeroNC.Length > 0)
			{
				MessageBox.Show(string.Format("Esta Factura esta Asociada con la Nota de Credito N. '{0}'", sNumeroNC), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
					
			if (((int)this.cmbFormaPago.Value == 6 || (int)this.cmbFormaPago.Value == 9) && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec VentaValidaPagos {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("Esta Factura ya tiene Pagos Registrados ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((int)this.cmbBodega1.Value == 36 || (int)this.cmbBodega1.Value == 81)
			{
				MessageBox.Show("No puede editar, Esta factura fue generada desde remates", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((int)this.txtNumEstado.Value != 6)
				if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			int idWeb = 0;
			idWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select count(*) from web_cabecera_pedido where idEstado_WEB > 2 and numero_pedido = {0}", 
				this.txtNumidOrigen.Value));

			#region Origen Factura 
			if ((int)this.txtNumidOrigen.Value > 0 && idWeb == 0) 
				idTipoOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idTipoFactura, 0) From Compra Where idCompra = {0}", 
					this.txtNumidOrigen.Value));

			if (idTipoOrigen == 0 || idTipoOrigen == 43)
			{
				this.txtSerial.Enabled = true;
				FuncionesProcedimientos.EstadoGrids(true, this.grdSerial);
			}
			
			if (idTipoOrigen == 5)
				FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			if ((int)this.txtNumidOrigen.Value == 0|| (int) this.cmbBodega1.Value == 82)
				if ((int)this.optTipoVenta.EditValue != 2)
					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);													
			#endregion Origen Factura 
			
			#region Permisos y Controles
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;						
			if ((int)this.txtNumEstado.Value == 6) this.btnTemporal.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnBuscar.Enabled = false;		

			this.txtNotas.Enabled = true;
			if (!bServicioTecnico && idTipoOrigen == 0 && bExigeVendedor) this.cmbVendedor.Enabled = true;
			this.txtNotas.Enabled = true;						
			this.chkEntregado.Enabled = true;	
			if((bool) this.chkEntregado.Checked) this.btnEntrega.Enabled = true;
			this.chkEntregaLocal.Enabled = true;
									
			this.txtSerial.Enabled = true;
			//Incremento por Angélica 
			this.txtresponsablecurier.Enabled = true;
			this.txtNumGuia.Enabled = true;
			this.txtValorEntrega.Enabled = true;
			this.txtNumTransaccion.Enabled = true;
			this.txtDireccionEntrega.Enabled = true;
			this.dtFechaEntregaCuriermbo1.Enabled = true;
			this.cmbCourier.Enabled = true;

			if (this.chkEntregado.Checked)
			{
				this.dtFechaEntrega.Enabled = true;
				this.dtpHoraEntrega.Enabled = true;
				this.cmbDirecciones.Enabled = true;
				this.txtNotas.Enabled = true;
				
			}

			if (miAcceso.CambiarDescuento && (int)this.txtNumEstado.Value == 6 && !bCrediPoint) this.btnClaveDescuento.Enabled = true;

			if (miAcceso.BEditarNombre)
			{
				if (!this.bCrediPoint)
				{
					this.cmbTipoRuc.Enabled = ((int)this.cmbFormaPago.Value != 35);
					this.txtRuc.Enabled = ((int)this.cmbFormaPago.Value != 35);
					this.btnCliente.Enabled = ((int)this.cmbFormaPago.Value != 35);
				}
			}

			this.btnValidarMail.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.CambiarNumeracion && (int)this.txtNumEstado.Value == 6) this.txtNumero.Enabled = true;
			//if (miAcceso.BModificarBodega) this.cmbBodega.Enabled = true; else this.cmbBodega.Enabled = false;
			if (miAcceso.BModifcarFormaDePago) 
				if ((int)this.cmbFormaPago.Value != 4 || (int)this.cmbFormaPago.Value != 6 || (int)this.cmbFormaPago.Value != 9)
					this.cmbFormaPago.Enabled = true;			
			#endregion Permisos y Controles
							
			this.ValidaAnticipos();					

			#region Retenciones
			if (BRetencion)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) == 0)
					EstadoRetencion(true);							
			}
			else
				if ((int)this.cmbTipoRuc.Value == 4 && ((int)this.cmbFormaPago.Value != 3 && (int)this.cmbFormaPago.Value != 11)) EstadoRetencion(true);						
			#endregion Retenciones
						
			#region Temporal
			if ((int)this.txtNumEstado.Value == 6)
			{
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico, Promociones From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
				dr.Read();
				if (dr.HasRows)
				{
					this.txtSerie.Text = dr.GetValue(0).ToString();
					this.txtAutFactura.Text = dr.GetValue(1).ToString();
					if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
					bNumeracionAutomatica = dr.GetBoolean(3);
					bPromociones = dr.GetBoolean(4);
				}
				dr.Close();

				if (!bNumeracionAutomatica)	this.txtNumero.Enabled = true; else this.txtNumero.Enabled = false;
				bNuevo = true;

				if (bPromociones)	
				{
					string[] sPrmCW = new string[2];
					sPrmCW = Facturacion.sEditarCuponWeb((int)this.txtNumIdCompra.Value, cdsSeteoF);
					idCuponWeb = int.Parse(sPrmCW[0]);
					
					if (idCuponWeb > 0) 
					{
						bCuponWeb = true;
						dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
					}
				}

				//if (idTipoOrigen == 0) this.optTipoVenta.EditValue = 1;
				this.optTipoVenta.Enabled = true;
			}
			else bNuevo = false;
			#endregion Temporal					

			bEdicion = true;		
	
			InhabilitarCedula();
		}

		private void NuevoRegistro(object sender, System.EventArgs e, int idFormaPago)
		{													
			bNuevo = true;
			bEdicion = true;

			#region Fechas
			this.txtNumContadoCredito.Value = 1;			
			this.optTipoVenta.EditValue = 1;						
			#endregion Fechas
									
			#region Habilita controles 
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;			
			/* BODEGA */
			
			if (!bServicioTecnico && bExigeVendedor) this.cmbVendedor.Enabled = true;
			this.txtRuc1.Enabled = false;

			if (bServicioTecnico)
			{
				this.cmbVendedor.Value = idTecnico;
				this.txtRuc.Text = sRuc;
				this.txtRuc1.Text = sRuc;
				CargaCliente();
			}

			if (idFormaPago != 4 && !bServicioTecnico)
			{
				this.cmbTipoRuc.Value = 5;	
				this.cmbTipoRuc.Enabled = (idFormaPago != 35);
				this.txtRuc.Enabled = (idFormaPago != 35);
				this.btnCliente.Enabled = (idFormaPago != 35);
			}

			if (bServicioTecnico)
			{
				this.ultraGrid1.Rows.Band.AddNew();
				int IPosicion = this.ultraGrid1.Rows.Count - 1;
				this.ultraGrid1.Rows[IPosicion].Cells["Codigo"].Value = sCodigoServTecn;
				this.ultraGrid1.Rows[IPosicion].Cells["Precio"].Value = dPrecio;
				this.ultraGrid1.Rows.Band.AddNew();
			}		

			if (!bFormaPagoPromocional) FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			FuncionesProcedimientos.EstadoGrids(true, this.grdSerial);
			this.grdSerial.Enabled = true;

			this.txtSerial.Enabled = true;
			this.optTipoVenta.Enabled = true;
			this.chkEntregado.Enabled = true;
			this.chkEntregaLocal.Enabled = true;
			this.txtNotas.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnValidarMail.Enabled = true;

			if (iEstadoDT == 0)
			{
				this.btnGuardar.Enabled = true;
				this.btnTemporal.Enabled = true;
			}
			else if (iEstadoDT == 6) this.btnTemporal.Enabled = true;
			else if (iEstadoDT == 10) this.btnGuardar.Enabled = true;

			this.btnCancelar.Enabled = true;
			//this.btnAnular.Enabled = true;
			this.btnCliente.Enabled = true;
			this.btnBuscar.Enabled = false;
			#endregion Habilita controles 									
			
			if (!bNumeracionAutomatica && bEdicion)	this.txtNumero.Enabled = true; else this.txtNumero.Enabled = false;

			if (miAcceso.CambiarFecha)
			{
				if (idFormaPago == 21)
				{
					this.dtFecha.Enabled = false; 
					this.txtNumero.Enabled = false;
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];	
				}
				else
				{
					this.dtFecha.Enabled = true; 
					this.dtFecha.Focus(); 
				}
			}
			else if ((int)this.txtBodegaPredef.Value > 0) this.cmbBodega.Focus();
			else if (bExigeVendedor) this.cmbVendedor.Focus();
			else this.txtRuc.Focus();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			FuncionesInventario._mostrarAsistenciaDental = true;
			this.btnCancelar_Click(sender, e);
			
			if (!bServicioTecnico)
			{
				if (bRebates)
				{
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
					bNuevo = true;

					this.cmbFormaPago.Value = 6;

					NuevoRegistro(sender, e, 6);
					return;
				}	
				using (frmFormasDePago FormaPago = new frmFormasDePago(1, miAcceso.BFormasDePagoCred, miAcceso.BFormasDePagoCE, (int)this.txtBodegaPredef.Value))
				{				
					if (DialogResult.OK == FormaPago.ShowDialog())
					{
						this.cmbFormaPago.Value = (int) FormaPago.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;
						bFormaPagoPromocional = (bool)FormaPago.ultraGrid1.ActiveRow.Cells["Promocional"].Value;
						
						FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
						bNuevo = true;

						#region Pago con nota de credito
						if ((int)this.cmbFormaPago.Value == 4)
						{						
							BUsoNotaDeCredito = true;
				
							this.txtNumIdCliente.Value = FormaPago.idCliente;
							this.cmbTipoRuc.Value = (int)FormaPago.idTipoRuc;
							this.txtRuc.Text = FormaPago.RUC.ToString();
							this.txtNombre.Text = FormaPago.Cliente.ToString();
														
							this.txtNumidOrigen.Value = (int)FormaPago.IdCompraNC;
							dSaldoNotaDeCredito = (decimal) FormaPago.dSaldo; 
							this.txtNumidFPNC.Value = FormaPago.IdFormaPagoNC;
							this.txtComprobante.Text = FormaPago.sNumero.ToString();						
							
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumidOrigen.Value);
							idBloqueaTransacciones1 = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

							this.ValidaAnticipos();
						}					
						#endregion Pago con nota de credito

						if ((int)this.txtBodegaPredef.Value == 0) 
						{
							this.cmbBodega.Enabled = true;
							this.cmbBodega.Value = 1; 				
						}
						else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

						#region Tarjeta de Credito
						if ((int)this.cmbFormaPago.Value == 3 || (int)this.cmbFormaPago.Value == 11)
						{
							using (frmInfoTarjetasDeCredito ITC = new frmInfoTarjetasDeCredito((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbFormaPago.Value))
							{				
								if (DialogResult.OK == ITC.ShowDialog())
								{					
									this.txtidBanco.Value = (int)ITC.cmbBanco.Value;
									this.txtidTarjeta.Value = (int)ITC.cmbTarjeta.Value;
									this.txtidPlan.Value = (int)ITC.cmbPlanes.Value;
									this.txtidPlazo.Value = (int)ITC.cmbPlazo.Value;							
								}
								else 
								{
									this.btnCancelar_Click(sender, e);
									this.cmbBodega.Value = System.DBNull.Value;
									return;
								}
							}
						}
						#endregion Tarjeta de Credito
						
						#region Forma de Pago Promocional
						if (bFormaPagoPromocional)
						{
							#region Valida Promocion
							int idPromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaPromocionPorFormaPagoPromocional ({0}, {1})", idTipoPrecio, (int)this.cmbBodega.Value));
								
							if (idPromocion > 0 && bPromociones)
							{	
								using (frmInformacionCupon IC = new frmInformacionCupon())
								{				
									if (DialogResult.OK == IC.ShowDialog())
									{
										#region Carga Informacion Articulos
										dtFecVencCupon = (DateTime)IC.dtFechaVencimiento.Value;
										sNumero = IC.txtNumero.Text.ToString();

										int iFila = 0;									
										string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, 1", idPromocion, 3, (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value);
										SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQLProm, true);

										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
	
										while (drC.Read())
										{
											bCargaPromocionFP = true;

											int idArtProm = drC.GetInt32(0);
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
												
										bCargaPromocionFP = false;

										this.Total();							
										#endregion Carga Informacion Articulos
									}
									else 
									{
										this.btnCancelar_Click(sender, e);
										return;
									}
								}
							}								
							#endregion Valida Promocion													
						}
						#endregion Forma de Pago Promocional		

						NuevoRegistro(sender, e, (int)this.cmbFormaPago.Value);						
					}
					else 
					{
						this.btnCancelar_Click(sender, e);
						return;
					}
				}
			}
			else
			{
				this.cmbFormaPago.Value = idFormaPagoST;
				NuevoRegistro(sender, e, (int)this.cmbFormaPago.Value);	
			}
		}
		
		bool bCargaPromocionFP = false;
		bool bFormaPagoPromocional = false;
		DateTime dtFecVencCupon = DateTime.Today;
		string sNumero = "";
		int idCre_SolicitudWeb = 0;

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{			
			if (this.cmbBodega.ActiveRow != null)
			{	
				if (bEdicion)
				{
					if ((int)this.cmbBodega.Value != 1 && (int)this.cmbBodega.Value != 101 && (int)this.cmbBodega.Value != 153)
					{
						int iTransferencias = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CompraTransf Where Generado = 1 And Bodega = {0}", (int)this.cmbBodega.Value));

						if (iTransferencias > 0)
						{
							MessageBox.Show(string.Format("No puede crear facturas hasta que procese la(s) {0} transferencias pendientes", iTransferencias), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.btnCancelar_Click(sender, e);
							return;
						}
					}
				}

				bUsoAnticipo = false;
				dValorAnticipo = 0.00m;
								
				if (bNuevo) this.txtNumero.Text = "";

				this.txtNumIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
				
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;

				this.cmbBodega1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 0", (int)this.cmbBodega.Value));
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 1, cdsSeteoF);

				if (bNuevo)
				{
					this.cmbBodega1.Value = System.DBNull.Value;
					if (this.cmbBodega1.Rows.Count == 1) 
						this.cmbBodega1.Value = Funcion.iEscalarSQL(cdsSeteoF, 
							string.Format("Select a.Bodega From AsignaBodegasLocales a Inner Join Bodega b On b.Bodega = a.Bodega Where Local = {0} And Remates = 0", 
							(int)this.cmbBodega.Value)); //Convert.ToInt32(this.cmbBodega.ActiveRow.Cells["Bodega"].Value);
					else this.cmbBodega1.Enabled = true;
				}

				if (this.cmbTipoRuc.ActiveRow == null) this.cmbTipoRuc.Value = 5;			
				
				if ((int)this.cmbBodega.Value == 69) if (bEdicion) this.txtClienteNombre.Enabled = true;
																						 else
																						 {
																							 if (bNuevo && (int)this.cmbFormaPago.Value != 4)
																							 {
																								 this.txtClienteNombre.Text = "";
																								 this.txtClienteNombre.Enabled = false;
																							 }
																						 }

				if (BUsoNotaDeCredito) idTipoPrecio = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.txtNumidFPNC.Value, (int)this.cmbBodega.Value));
				else idTipoPrecio = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value));
				
				#region S.R.I. si es registro nuevo
				if (bNuevo)
				{
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico, Promociones, ValidaTipoVenta From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
					dr.Read();
					if (dr.HasRows)
					{
						this.txtSerie.Text = dr.GetValue(0).ToString();
						this.txtAutFactura.Text = dr.GetValue(1).ToString();
						if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
						bNumeracionAutomatica = dr.GetBoolean(3);
						bPromociones = dr.GetBoolean(4);
						bValidaTipoVenta = dr.GetBoolean(5);
					}
					dr.Close();	
					
					if (!bNumeracionAutomatica && bEdicion)	
						this.txtNumero.Enabled = true; 
					else
					{
						this.txtNumero.Text = "";
						this.txtNumero.Enabled = false;
					}
			
					if ((int)this.txtNumIdCompra.Value == 0 && bPromociones)	
					{
						/*
						string[] sPrmCW = new string[2];
						sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
						idCuponWeb = int.Parse(sPrmCW[0]);
            					
						if (idCuponWeb > 0)
						{
							bCuponWeb = true;
							dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
						}
						*/
						//						{
						//							bCargaPromocionFP = true;
						//							FuncionesInventario.CargaArticulosPromocionCupon(this.ultraGrid1, (DateTime)this.dtFecha.Value, 1, (int)this.cmbBodega.Value, 0, (int)this.cmbFormaPago.Value, idTipoPrecio, 
						//								(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, -1, 0, cdsSeteoF);
						//							bCargaPromocionFP = false;
						//							this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
						//							Total();
						//							this.cmbVendedor.Focus();
						//						}
					}
				}
				#endregion S.R.I. si es registro nuevo				
				#region Integración Yupana (Forma de Pago 35)
				/*if (bNuevo && this.cmbFormaPago.Value != null && this.cmbFormaPago.Value != DBNull.Value && (int)this.cmbFormaPago.Value == 35)
				{
					bool bRegistroSeleccionado = false;
					using (frmBuscaCre_SolicitudYupana frmYupana = new frmBuscaCre_SolicitudYupana((int)this.cmbBodega.Value))
					{
						if (frmYupana.ShowDialog() == DialogResult.OK && frmYupana.ultraGrid1.ActiveRow != null)
						{
							if (frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value != null && frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value != DBNull.Value)
							{
								this.cmbTipoRuc.Value = 5;
								this.txtRuc.Text = frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value.ToString();
								this.txtRuc_Validated(this.txtRuc, System.EventArgs.Empty);
								this.cmbNumContadoCredito.Enabled = true;
								bRegistroSeleccionado = true;
							}
						}
					}

					if (!bRegistroSeleccionado)
					{
						MessageBox.Show("En la bodega seleccionada no existe registro de solicitud Yupana", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.cmbBodega.Value = System.DBNull.Value;
					}
				}*/
				#endregion
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{			
			if (bEdicion)
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 1, (int)this.cmbBodega.Value);
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
				try
				{
					if (!bEliminaAlValidar)
					{
						if (!FuncionesInventario.ValidaEliminaFilasGrids(sender, e, 1, this.grdSerial, cdsSeteoF)) return;

						if (!bFormaPagoPromocional)
						{						
							#region Validacion
							if ((int)this.cmbFormaPago.Value == 9)
							{
								MessageBox.Show("No puede eliminar filas de Facturas de CREDIPOINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}

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

							if (!bNuevo && !miAcceso.BEliArt) 
							{
								MessageBox.Show("No puede eliminar filas de Facturas ya registradas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}
							#endregion Validacion			

							if (!bEliminando) 
							{
								bEliminando = true;				
								if (!Facturacion.FacturacionEliminaFilas(this.ultraGrid1, this.grdSerial, e, this.btnCancelar, 43, bNuevo, bImpreso, miAcceso.BEliArtCmbProm, miAcceso.BEliArt, (int)this.cmbBodega.Value, idTipoPrecio, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value,
									(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlazo.Value, (int)this.txtidPlan.Value, 0, bCuponWeb, false, cdsSeteoF)) e.Cancel = true;
								e.DisplayPromptMsg = false;
								bEliminando = false;						
							}
							else e.DisplayPromptMsg = false;
						}
					}
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Error al eliminar filas\n\nSi la Factura esta en estado temporal se anulara automaticamente\n\nLa Pantalla se cerrara vuelva a abrirla para continuar", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					if ((int)this.txtNumIdCompra.Value > 0 && (int)this.txtNumEstado.Value == 6) 
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulacionCompraVenta {0}, 1, 'ANULACION AUTOMATICA POR ERROR EN ELIMINACION DE PRODUCTOS'", (int)this.txtNumIdCompra.Value));

					this.btnCancelar_Click(sender, e);
					this.Close();
					e.DisplayPromptMsg = false;
				}
			}
		}

		bool bEliminando = false;
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
			/* */
			if (e.Row.Cells["Estado"].Value != null && Convert.ToString(e.Row.Cells["Estado"].Value)!= "" )
			{
				if ((int)e.Row.Cells["Estado"].Value == 11)
				{
					e.Row.Appearance.BackColor = Color.Red;
					e.Row.Appearance.BackColor2 = Color.Red;
				}
			}
		}

		


	
		private void grdSerial_DoubleClick(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.grdSerial.ActiveRow.Cells["idSerial"].Value != System.DBNull.Value) 
				{
					if (idTipoOrigen == 3 || idTipoOrigen == 11 || idTipoOrigen == 13) return;

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

				idAgente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAgenteRetencion, 0) From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value));
				idSujeto = Funcion.iEscalarSQL(cdsSeteoF, "Select ISNULL(idAgenteRetencion, 0) From Seteo");
						
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
							dr1.Cells["Base"].Value = Math.Round(Convert.ToDecimal(dr1.Cells["Base"].Value), 2) + Math.Round(dBaseIVA, 2);
							dr1.Cells["Valor"].Value = Math.Round(Convert.ToDecimal(dr1.Cells["Valor"].Value), 2) + Math.Round(dIVARetenido, 2);
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
							dr1.Cells["Base"].Value = Math.Round(Convert.ToDecimal(dr1.Cells["Base"].Value), 2) + Math.Round(dBaseRenta, 2);
							dr1.Cells["Valor"].Value = Math.Round(Convert.ToDecimal(dr1.Cells["Valor"].Value), 2) + Math.Round(dRentaRetenida, 2);
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
					dValorRetencion = dValorRetencion + Math.Round(Convert.ToDecimal(dr.Cells["Valor"].Value), 2);
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
			if (bEdicion)
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;
		}
		private void txtRuc1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			//		if (bEdicion)
			//				if (!Validacion.vbIdentificacion(this.txtRuc1, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;	
		}

		private void ValidaAnticipos()
		{
			bUsoAnticipo = false;
			dAnticipo = 0;

			if (((int)this.cmbFormaPago.Value == 9 && (decimal)this.txtCuotaDeEntrada.Value == 0)) return;
			
			dAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int) this.txtNumIdCliente.Value));
				
			if (dAnticipo > 0)
			{
				if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 				
					bUsoAnticipo = true;
				else
				{
					bUsoAnticipo = false;
					dAnticipo = 0.00m;
				}
			}									
		}

		int iEncuesta = 0;
		string sOtrosEncuesta = "";

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Cliente
				if ((int)this.cmbTipoRuc.Value != 11)
				{
					this.txtNumIdCliente.Value = 0;
					this.txtNombre.Text = "";

					//					CargaCliente();

					//					if ((int)this.txtNumIdCliente.Value == 0)
					//					{					
					FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), false, miAcceso.EditarCuentaNombre);
					CargaCliente();			
	
					bEmpleado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaEmpleadoActivo '{0}'", this.txtRuc.Text.ToString()), true);

					#region Encuesta
					if (bNuevo)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, "Select Encuesta From SeteoF", true))
						{
							using (frmFacturacionEncuesta FE = new frmFacturacionEncuesta())
							{
								if (DialogResult.OK == FE.ShowDialog())
								{
									iEncuesta = (int)FE.optEncuesta.Value;
									sOtrosEncuesta = FE.txtOtros.Text.ToString();
								}
								else return;
							}
						}
					}
					#endregion Encuesta

					//					}
				
					this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtNumIdCliente.Value, cdsSeteoF);
				
					ValidaAnticipos();

					if ((int)this.cmbTipoRuc.Value == 4 && ((int)this.cmbFormaPago.Value != 3 && (int)this.cmbFormaPago.Value != 11)) EstadoRetencion(true); else EstadoRetencion(false);
					

					#region FamilyFriends
					VerificarFamilyFriends();
					#endregion
					InhabilitarCedula();
					#endregion Busca Cliente


				}
			}
		}

		private void InhabilitarCedula()
		{
			if(this.txtRuc.Text != null && this.txtRuc.Text != "")
			{
				this.txtRuc.Enabled = false;
				this.btnCliente.Enabled = false;
				this.cmbTipoRuc.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbFormaPago.Enabled = false;
			}
		}
		

		// Request que recuerda la última cédula y forma de pago confirmadas
		private FamilyFriendsRequestModel reqFF = new FamilyFriendsRequestModel();
		// Solo una variable para todo el estado FF
		private FamilyFriendsEstadoModel estadoFF = new FamilyFriendsEstadoModel();

		private void VerificarFamilyFriends()
		{
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			// Reutilizamos el mismo request global
			reqFF.ValorFormaPago = this.cmbFormaPago.Value;
			reqFF.RucCliente = this.txtRuc.Text;
			reqFF.Owner = this;

			// Ejecuta la verificación
			FamilyFriendsVerificacionModel r = service.VerificarFF(reqFF);

			// Guardar en el modelo
			estadoFF.EsFF = r.EsFF;
			estadoFF.IdCliente = r.IdCliente;
			estadoFF.IdWEB_FF_Invitado = r.IdWEB_FF_Invitado;
			estadoFF.RechazoPromocion = r.RechazoPromocion;

			// ==============================
			// Manejo de estados FF
			// ==============================
			if (estadoFF.EsFF)
			{
				idTipoPrecio = r.IdTipoPrecio;
				this.lblRecurrente.Text = "CLIENTE FAMILY & FRIENDS";
			}
			else
			{
				estadoFF.IdCliente = 0;

				if (estadoFF.RechazoPromocion)
				{
					this.lblRecurrente.Text = "PROMOCIÓN FF RECHAZADA";
				}
				else
				{
					this.lblRecurrente.Text = "";
				}
			}

			// Mostrar mensaje del SP si existe
			if (r.MensajeSP != null && r.MensajeSP.Length > 0)
			{
				MessageBox.Show(
					r.MensajeSP,
					"Family & Friends",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information
					);
			}
		}

		private void txtRuc1_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Cliente 
				this.txtNumIdCliente1.Value = 0;
				this.txtClienteNombre.Text = "";

				CargaCliente();

				if ((int)this.txtNumIdCliente1.Value == 0)
				{				
					FuncionesProcedimientos.AbreFacturacionClientes1(false,0,(int)this.txtNumIdCliente1.Value);						
					CargaCliente();					
				}				
				this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtRuc1.Value, cdsSeteoF);
				
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
				this.mnuImprimirFactura.Enabled=false;
				this.mnuImprimirContrato.Enabled=false;
				this.mnuImprimirNuovo.Enabled=false;
				this.mnuDocsCrediPoint.Enabled = false;
				this.mnuDocsCrediPointBlanco.Enabled = false;
				
				this.mnuImprimirGarantiaExtendida.Enabled = false;
				this.mnuImprimirDYD.Enabled = false;
				this.mnuImprimirGEP.Enabled = false;
				this.mnuImprimirCP.Enabled = false;
				this.mnuImprimirPR.Enabled = false;
				this.mnuImprimirASD.Enabled = false;
				#endregion Botones

				#region Maestro
				this.txtNumIdCompra.Value = IdCompra;
				int iBodega = 0;
				int iBodegaInv = 0;
				int idVendedor = 0;
				
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
					idVendedor = dr.GetInt32(12);
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
					this.txtProyecto.Text = dr.GetString(45);
					
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
					bVeriricador = dr.GetBoolean(56);
					this.optTipoVenta.EditValue = dr.GetInt32(57);
					this.txtNumTransaccion.Text = dr.GetString(58);
					this.txtDireccionEntrega.Text = dr.GetString(59);
					this.cmbCourier.Value = dr.GetInt32(60); 
					this.txtNumGuia.Text = dr.GetString(61);
					this.txtValorEntrega.Value = dr.GetDecimal(62);
					this.txtClienteNombre.Text = dr.GetString(63);
					iBodegaInv = dr.GetInt32(64);
					this.txtresponsablecurier.Text = dr.GetString(65);
					if (dr.GetValue(66) != System.DBNull.Value) this.dtFechaEntregaCuriermbo1.Value = dr.GetDateTime(66);
					if (dr.GetValue(67) != System.DBNull.Value) idFe_EstadoDocumento = dr.GetInt32(67); 

					//Cargar el valor de la encuesta
					iEncuesta = dr.GetInt32(68);
					sOtrosEncuesta = dr.GetString(69);
          
					this.chkEntregaLocal.Checked = dr.GetBoolean(70);
					CodigoEntrega = dr.GetString(71);
					idCre_SolicitudWeb  = dr.GetInt32(72);
				}
				dr.Close();

				this.txtBaseImponible.Value = (decimal)this.txtIva.Value - (decimal)this.txtDescIvaTotal.Value;

				this.cmbBodega.Value = iBodega;
				this.cmbFechaPrimerPago.Value = dtFechPrimPago;
				this.cmbNumContadoCredito.Value = NumCuotas;

				this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtNumIdCliente.Value, cdsSeteoF);
        if((int) this.cmbFormaPago.Value == 35 && idCre_SolicitudWeb > 0 ) this.lblYupana.Text ="YUPANA";
				if ((int)this.txtNumEstado.Value == 6 && (int)this.cmbFormaPago.Value == 4) 
				{
					BUsoNotaDeCredito = true;
					dSaldoNotaDeCredito = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(SaldoAnticipo, 0) From Compra Where idCompra = {0} And idCliente = {1} And idTipoFactura = 5", (int)this.txtNumidOrigen.Value, (int)this.txtNumIdCliente.Value));
				}
				#endregion Maestro

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value));
				this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value));

				if ((int)this.cmbFormaPago.Value == 3) this.lblPlan.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select Nombre from banco where idBanco = {0}", (int)this.txtidBanco.Value));	 
				if ((int)this.cmbFormaPago.Value == 3) this.lblPlan.Text = this.lblPlan.Text + " "+ (Funcion.sEscalarSQL(cdsSeteoF, string.Format("select Descripcion from Plazos where idPlazo= {0}", (int)this.txtidPlazo.Value)));				
				if ((int)this.cmbFormaPago.Value == 3) this.lblPlan.Text = this.lblPlan.Text + " "+(Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Descripcion From Planes Where idPlan = {0}", (int)this.txtidPlan.Value)));		
				
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

				#region Pago con nota de credito
				//				if ((int)this.cmbFormaPago.Value == 4)
				//				{						
				//					BUsoNotaDeCredito = true;
				//					
				//					string sSQLNC = string.Format("Select (Select idFormaPago From Compra where idCompra = c.idOrigen), SaldoAnticipo, Numero From Compra c Where idCompra = {0}", (int)this.txtNumidOrigen.Value);
				//					SqlDataReader drNC = Funcion.rEscalarSQL(cdsSeteoF, sSQLNC, true);
				//					drNC.Read();
				//					if (drNC.HasRows)
				//					{
				//						this.txtNumidFPNC.Value = drNC.GetInt32(0);
				//						dSaldoNotaDeCredito = drNC.GetDecimal(1);						
				//						this.txtComprobante.Text = drNC.GetString(2);
				//					}
				//					drNC.Close();					
				//					
				//					idTipoPrecio = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTipoDePrecioVenta {0}, {1}", (int)this.txtNumidFPNC.Value, (int)this.cmbBodega.Value));					
				//
				//					idCompraOrigen = (int)this.txtNumidOrigen.Value;
				//
				//					string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumidOrigen.Value);
				//					idBloqueaTransacciones1 = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
				//				}					
				#endregion Pago con nota de credito

				this.cmbBodega1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 0", (int)this.cmbBodega.Value));
				this.cmbBodega1.Value = iBodegaInv;
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbVendedor.Value = idVendedor;
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 1, cdsSeteoF);

				#region Habilita Controles - permisos de usuario
				if (!bAnulado)
				{
					if ((int)this.txtNumEstado.Value != 6)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
						{
							if (miAcceso.Editar) this.btnEditar.Enabled = true;// && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) 

							if (miAcceso.Anular && !bAnulado && (int)this.cmbFormaPago.Value != 9) this.btnAnular.Enabled = true; //ANULA FACTURAS QUE NO SON CREDIPOINT
							if (miAcceso.AnularCrediPoint && !bAnulado && (int)this.cmbFormaPago.Value == 9) this.btnAnular.Enabled = true; //ANULA FACTURAS QUE SON CREDIPOINT

							if (this.btnAnular.Enabled && miAcceso.BAnulaSoloDelDia && !bAnulado)
								if ((DateTime)this.dtFecha.Value != DateTime.Today) this.btnAnular.Enabled = false;

							if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							if (this.chkEntregado.Checked) this.btnImprimeED.Enabled = true;
						}
					}
					else
					{
						if (miAcceso.Editar && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) this.btnEditar.Enabled = true;

						if (miAcceso.Anular && !bAnulado && (int)this.cmbFormaPago.Value != 9) this.btnAnular.Enabled = true; //ANULA FACTURAS QUE NO SON CREDIPOINT
						if (miAcceso.AnularCrediPoint && !bAnulado && (int)this.cmbFormaPago.Value == 9) this.btnAnular.Enabled = true; //ANULA FACTURAS QUE SON CREDIPOINT

						if (this.btnAnular.Enabled && miAcceso.BAnulaSoloDelDia && !bAnulado)
							if ((DateTime)this.dtFecha.Value != DateTime.Today) this.btnAnular.Enabled = false;

						if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
						//if (miAcceso.CrearAsiento && !bAnulado) this.btnAsiento.Enabled = true;
					}

					#region Impresion
					if ((int)this.txtNumEstado.Value == 10)
					{
						if (!bImpreso) 
						{
							if (miAcceso.Imprimir) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirFactura.Enabled = true;	
								this.mnuImprimirContrato.Enabled = true;
								this.mnuImprimirNuovo.Enabled=true;
							}

							int iContGE = 0;
							int iContDYD = 0;
							int iContGEP = 0;
							int iContCP = 0;
							int iContPR = 0;
							int iContASD = 0;

							#region Validación Error PayJoy
							bool bErrorPayJoy = false;
							int idArticuloError = 0;
							int idFormaPagoCrerditoPayJoy = 0;
							string cedulaPayJoy = "";
							int idPayJoyConfiguracionErrorPayJoy = 1;
							string Mensaje = "";

							bool Estado = false;

							try
							{
								string squery = String.Format("Exec PayJoyConfiguracionModel '{0}', {1}", 
									"Estado", 
									idPayJoyConfiguracionErrorPayJoy);
								Estado = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, squery, false);
								squery = String.Format("Exec PayJoyConfiguracionModel '{0}', {1}", 
									"idFormaPagoCrerditoPayJoy", 
									idPayJoyConfiguracionErrorPayJoy);
								idFormaPagoCrerditoPayJoy = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery,false);
								squery = String.Format("Exec PayJoyConfiguracionModel '{0}', {1}", 
									"cedulaPayJoy", 
									idPayJoyConfiguracionErrorPayJoy);
								cedulaPayJoy = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
								squery = String.Format("Exec PayJoyConfiguracionModel '{0}', {1}", 
									"Mensaje", 
									idPayJoyConfiguracionErrorPayJoy);
								Mensaje = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
							}
							catch(Exception ex)
							{
								string sms = ex.Message;
							}
							#endregion Validación Error PayJoy

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drGE in ultraGrid1.Rows.All)
							{
								#region Validación Error PayJoy
								if(Estado)
								{
									try
									{
										idArticuloError = (int)drGE.Cells["idArticulo"].Value;
									}
									catch(Exception ex)
									{
										string sms = ex.Message;
										bErrorPayJoy = true;
									}
									if((int)this.cmbFormaPago.Value == idFormaPagoCrerditoPayJoy && bErrorPayJoy && cedulaPayJoy == this.txtRuc.Text)
									{
										MessageBox.Show(string.Format(Mensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										return;
									}
								}
								#endregion Validación Error PayJoy

								if ((int)drGE.Cells["Tipo"].Value == 5) iContGE++;
								if ((int)drGE.Cells["idArticulo"].Value == 8683) iContDYD++;
								if ((int)drGE.Cells["Tipo"].Value == 9) iContGEP++;
								if ((int)drGE.Cells["Tipo"].Value == 10) iContCP++;
								if ((int)drGE.Cells["idArticulo"].Value == 12876) iContPR++;
								if ((int)drGE.Cells["idArticulo"].Value == 12929 ||(int)drGE.Cells["idArticulo"].Value == 12930) iContASD++;
								//if ((int)drGE.Cells["idArticulo"].Value == 12920 ||(int)drGE.Cells["idArticulo"].Value == 12921) iContASD++;
							}

							if (iContGE > 0) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirGarantiaExtendida.Enabled = true;
							}						
  	
							if (iContDYD > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirDYD.Enabled = true;
							}

							if (iContGEP > 0) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirGEP.Enabled = true;
							}						
  	
							if (iContCP > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirCP.Enabled = true;
							}
							
							if (iContPR > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirPR.Enabled = true;
							}

							if (iContASD > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirASD.Enabled = true;
							}
						}

						if (bImpreso) 
						{
							if (miAcceso.ReImprimir) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirFactura.Enabled = true;
								this.mnuImprimirContrato.Enabled = true;
								this.mnuImprimirNuovo.Enabled=true;
							}

							int iContGE = 0;
							int iContDYD = 0;
							int iContGEP = 0;
							int iContCP = 0;
							int iContPR = 0;
							int iContASD = 0;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drGE in ultraGrid1.Rows.All)
							{
								if ((int)drGE.Cells["Tipo"].Value == 5) iContGE++;
								if ((int)drGE.Cells["idArticulo"].Value == 8683) iContDYD++;
								if ((int)drGE.Cells["Tipo"].Value == 9) iContGEP++;
								if ((int)drGE.Cells["Tipo"].Value == 10) iContCP++;
								if ((int)drGE.Cells["idArticulo"].Value == 12876) iContPR++;
								if ((int)drGE.Cells["idArticulo"].Value == 12929 || (int)drGE.Cells["idArticulo"].Value == 12930) iContASD++;
								//if ((int)drGE.Cells["idArticulo"].Value == 12920 ||(int)drGE.Cells["idArticulo"].Value == 12921) iContASD++;
							}
							
							if (iContGE > 0) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirGarantiaExtendida.Enabled = true;
							}						
  	
							if (iContDYD > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirDYD.Enabled = true;
							}

							if (iContGEP > 0) 
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirGEP.Enabled = true;
							}						
  	
							if (iContCP > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirCP.Enabled = true;
							}

							if (iContPR > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirPR.Enabled = true;
							}

							if (iContASD > 0)
							{
								this.btnImprimir.Enabled = true;
								this.mnuImprimirASD.Enabled = true;
							}

						}
					
						if ((int)this.cmbFormaPago.Value == 9) 
						{
							this.btnImprimir.Enabled = true;
							this.mnuDocsCrediPoint.Enabled = true;
							this.mnuDocsCrediPointBlanco.Enabled = true;							
						}
	
						if (idFe_EstadoDocumento == 2)
						{
							this.btnImprimir.Enabled = true;
							this.mnuImprimirRIDE.Enabled = true;							
						}
					}
					#endregion Impresion

					if (miAcceso.BBuscarDocumentos)
						this.btnBuscar.Enabled = true;
				}
				else if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;
				#endregion Habilita Controles - permisos de usuario
				        
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

			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value, miAcceso.BLimiteBusquedaFechas, miAcceso.BDia))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
					#region FamilyFriends
					ConsultarEstadoFamilyFriends((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
					#endregion FamilyFriends
				}
			}		
		}

		private void ConsultarEstadoFamilyFriends(int idCompra)
		{
			try
			{
				FamilyFriendsService service = new FamilyFriendsService(frmRequerimiento.sconexionPoint);

				// Obtener temporal para validar SIN tocar lo actual
				FamilyFriendsEstadoModel temp = service.ObtenerEstadoFF(idCompra);

				// Solo actualizar si SÍ hay registro válido
				if (temp.EsFF || temp.RechazoPromocion)
				{
					estadoFF = temp; // Actualiza el modelo principal
					idTipoPrecio = estadoFF.IdTipoPrecio;
					this.lblRecurrente.Text = estadoFF.MensajeEstadoFF;
				}
				else
				{
					// NO tocar estadoFF ni idTipoPrecio
					this.lblRecurrente.Text = "";
				}
			}
			catch
			{
				// No cambiar nada si hubo error
				this.lblRecurrente.Text = "";
			}
		}


		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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
				return;
			}
			
			#region Nuevo Registro	
			if (bNuevo)
			{					
				if (this.txtNumero.Text.ToString().Length == 0 && !bNumeracionAutomatica)
				{
					MessageBox.Show("Ingrese el número de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();						
					return;
				}
					
				string sSQLVNE = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)this.txtNumIdCompra.Value, this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVNE, true);

				if (sMensaje.ToString().Length > 0 && !bNumeracionAutomatica)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();					
					return;
				}
										
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{
						#region Maestro
						this.lblEstado.Text = "ANULADO";
						bAnulado = true;
						int idDireccion = 0;
						int idVendedor = 0;
						this.txtNumEstado.Value = 10; 
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

						string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, 0.0, 0.0, 0.0, 0.0, 0, 0.0, 0, 0, 0, 0, 0, '', 0, 1, '', '', 0, '', 0", 
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
						this.txtRuc1.Enabled = false;
						this.dtFecha.Enabled = false;

						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
						if (!miAcceso.Eliminar) this.btnAnular.Enabled = false;
						this.btnGuardar.Enabled = false;
						this.btnTemporal.Enabled = false;
						this.btnAnular.Enabled = false;
						this.btnCancelar.Enabled = true;
						this.btnClaveDescuento.Enabled = false;
						this.btnCliente.Enabled = false;
						bDescuentoActivo = false;

						bNuevo = false;
						bEdicion = false;
									
						HabilitaTab(false);
						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
						FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);
						#endregion Controles

						#region Impresión
						if (bNuevo)
						{
							frmMensajeNumeración miMsje = new frmMensajeNumeración (this.txtNumero.Text.ToString(), bAnulado, 2, false);
						}
						#endregion Impresión									
					}
				}											
				return;
			}
			#endregion Nuevo Registro
				
			this.Consultar((int)this.txtNumIdCompra.Value);

			#region Elimina Registro
			if (bAnulado)
			{						
				if (!miAcceso.Eliminar)
				{
					MessageBox.Show("El Documento esta Anulado, No tiene acceso a Eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
					return;
				}
				if (DialogResult.Yes == MessageBox.Show("Esta seguro de ELIMINAR este Documento \n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					string sSQLBorrar = string.Format("Exec BorrarCompraVenta {0}, {1}, {2}", (int)this.txtNumIdCompra.Value, (int)this.txtNumIdAsiento.Value, (int)this.txtIdTipoFactura.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBorrar, true);

					MessageBox.Show("Documento Eliminado del sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnCancelar_Click(sender, e);							
				}
				return;						
			}
			#endregion Elimina Registro

			#region Valida Anulacion
			if (bAnulado)
			{
				MessageBox.Show("El Documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}
			if (!miAcceso.Anular && (int)this.cmbFormaPago.Value != 9)
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}
			if (!miAcceso.AnularCrediPoint && (int)this.cmbFormaPago.Value == 9)
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim() + " de CrediPOINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}		
   
			if(!miAcceso.AnularDocSRI)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFE_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) == 2)
				{
					MessageBox.Show("No puede Anular la Factura esta autorizada por el S.R.I.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Bloquear, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true))
			{
				MessageBox.Show("No puede Anular la Factura esta Bloqueada", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}							
			string sNumeroNC = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VentasValidaNotaDeCredito {0}", (int)this.txtNumIdCompra.Value));
			if (sNumeroNC.Length > 0)
			{
				MessageBox.Show(string.Format("Esta Factura esta Asociada con la Nota de Credito N. '{0}'", sNumeroNC), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesVenta Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("No puede Anular la Factura esta Asociada a una Retencion que a Afecta a Multiples Facturas, \n\nAnule la Retencion o Consulte al Administrador", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if (((int)this.cmbFormaPago.Value == 6 || (int)this.cmbFormaPago.Value == 9) && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec VentaValidaPagos {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("Esta Factura ya tiene Pagos Registrados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!FuncionesProcedimientos.CVValidaCapturadoDepositado ((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;
			if (!FuncionesInventario.ValidaBodegaInventarios(this.ultraGrid1, (int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;
			if (Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;

			if ((int)this.cmbFormaPago.Value != 6)
				if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;
		
			DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
			if ((miAcceso.BAnulaSoloDelDia && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Estado, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) != 6) && dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
			{
				MessageBox.Show("Solo puede Anular Facturas de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}

			if (miAcceso.BAnulaSoloDelDia && dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString())
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetalleFormaPago Where Origen = 1 And idFormaPago = 1 And idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
				{
					MessageBox.Show("No tiene acceso a Anular facturas con pagos en efectivo " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);                                                                                       
					return;
				}
			}
			#endregion Valida Anulacion

			#region Anulacion	
			if (Facturacion.AnulacionCompraVenta ((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))
			{
				if (idCuponWeb > 0) 
					Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, 0, cdsSeteoF);

				this.Consultar((int)this.txtNumIdCompra.Value);
			}
			#endregion Anulacion		
			
			#region Bono Entrada 25
			int idBonoEntradaTipo = 0;
			//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
			DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
			if (dvResultado != null)
			{
				foreach (DataRowView rowView in dvResultado)
				{
					DataRow row = rowView.Row;
					idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
					//Validar si el idOrigen, pertenecese a la tabla BonoEntrada, con Bono 25
					int idBonoEntradaIdPrefactura = 0;
					string squery = String.Format("Exec BonoEntradaConsultarIdPrefactura {0}", (int)this.txtNumIdCompra.Value);
					idBonoEntradaIdPrefactura = Funcion.iEscalarSQL(cdsSeteoF, squery);
					if(idBonoEntradaIdPrefactura > 0)
					{
						//Bono entrada 25 activar
						frmPrefactura.BonoEntradaInactivar(cdsSeteoF,idBonoEntradaIdPrefactura, idBonoEntradaTipo, 1);
						frmPrefactura.BonoEntradaInactivarRelacion(cdsSeteoF,idBonoEntradaIdPrefactura, idBonoEntradaTipo, 1);
					}
				}
			}
			#endregion Bono Entrada 25
		}	

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + this.txtNumIdCompra.Value.ToString();
			
			stReporte += ", 0";
						
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);

			//stReporte = "VentasNotasDeEntregaFP.rpt";

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
				if (iTipo == -2) // DENTALIS
				{
					ParameterFields paramFields = new ParameterFields ();				
			
					ParameterField paramFieldEmpresa = new ParameterField ();
					ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
					paramFieldEmpresa.ParameterFieldName = "@idCompra";
					discreteValEmpresa.Value = (int)this.txtNumIdCompra.Value; 
					paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
					paramFields.Add (paramFieldEmpresa);

					Reporte miRep = new Reporte("FacturaDentalis.rpt", "");
					miRep.MdiParent = this.MdiParent;
					miRep.Titulo("Transferencia");
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.Show();
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			//string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + this.txtNumIdCompra.Value.ToString();
			//Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			bImpreso = true;
			if (!miAcceso.ReImprimir) this.btnImprimir.Enabled = false;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo			

			#region Imprime Celulares 
			if ((int)this.cmbFormaPago.Value==9 && (int)this.txtIdEF.Value == 14)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idCompra";
				ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				Reporte miRep = new Reporte("ContratoCompraVentaCelular.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Recepción Mercaderia de Retiro");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();	
			}
			#endregion Imprime Celulares 
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 20));
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) e.Cancel = true;
		}

		private void mnuBuscarPrefactura_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.mnuBI.Show(this.btnBuscar, new Point(1, 20));			
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbVendedor.Focus();
		}
		
		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (bEliminaAlValidar) return;

				FuncionesInventario.ActualizaPosicionArticulosSeriales(e, this.ultraGrid1, this.grdSerial);
		
				if (this.ultraGrid1.Rows.Count == 0) if ((int)this.cmbBodega.Value == 69) this.txtClienteNombre.Enabled = true;
				Total();
			}
		}
		
		private void btnCliente_Click(object sender, System.EventArgs e)
		{				
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
						
			FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, miAcceso.EditarCuentaNombre);
				
			CargaCliente();
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
						this.dtFecha.Value = DateTime.Today;
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
						this.txtSerie.Text = miBusqueda.grdFacturas.ActiveRow.Cells["SerieFactura"].Value.ToString();
						this.txtAutFactura.Text = miBusqueda.grdFacturas.ActiveRow.Cells["AutFactura"].Value.ToString();
						if (miBusqueda.grdFacturas.ActiveRow.Cells["FechaCaducidad"].Value != System.DBNull.Value) this.cmbCaducidad.Value = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["FechaCaducidad"].Value;
						if ((bool)miBusqueda.grdFacturas.ActiveRow.Cells["Entregado"].Value)
						{
							this.chkEntregado.Checked = (bool)miBusqueda.grdFacturas.ActiveRow.Cells["Entregado"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["FechaEntrega"].Value != System.DBNull.Value) this.dtFechaEntrega.Value = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["FechaEntrega"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["HoraEntrega"].Value != System.DBNull.Value) this.dtpHoraEntrega.Value = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["HoraEntrega"].Value;
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
						this.txtNumIdproyectoZoom.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idProyectoZOOM"].Value;
						bImpreso = false;
						this.txtNumIdAsiento.Value = 0;
						bAnulado = false;
						#endregion Maestro

						#region S.R.I. 
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico, Promociones From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
						dr.Read();
						if (dr.HasRows)
						{
							this.txtSerie.Text = dr.GetValue(0).ToString();
							this.txtAutFactura.Text = dr.GetValue(1).ToString();
							if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
							bNumeracionAutomatica = dr.GetBoolean(3);
							bPromociones = dr.GetBoolean(4);
						}
						dr.Close();	
			
						if (!bNumeracionAutomatica)	this.txtNumero.Enabled = true;
						#endregion S.R.I.

						this.cmbBodega1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}", (int)this.cmbBodega.Value));
						this.cmbBodega1.Value = (int)this.cmbBodega.Value;

						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumidOrigen.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo);

						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePedidosReservasConsignaciones {0}", (int)this.txtNumidOrigen.Value));	

						this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec BuscaSerialesPedido {0}", (int)this.txtNumidOrigen.Value));
						
						bClienteFacturaBajoCosto = 
							Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FacturaBajoCosto From Cliente Where idCliente = {0}", 
							(int)this.txtNumIdCliente.Value), false);
			
						this.ValidaAnticipos();
						
						this.optTipoVenta.Enabled = true;
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

						this.dtFecha.CalendarInfo.MinDate = (DateTime)miBusqueda.grdFacturas.ActiveRow.Cells["Fecha"].Value;
						this.optTipoVenta.EditValue = 1;

						this.btnNuevo.Enabled = false;
						this.btnConsultar.Enabled = false;						
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

		private void btnTemporal_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "VENTA", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (bExigeVendedor)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;
						
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;			
			
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, (int)this.cmbFormaPago.Value, cdsSeteoF)) return;
			
			if (!Validacion.EntregaDomicilio(this.chkEntregado, this.cmbDirecciones, this.dtFechaEntrega, this.dtpHoraEntrega, (int)this.optTipoVenta.EditValue)) return;

			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 1)) return;

			if ((bool) this.chkEntregado.Checked )
			{
				if( CodigoEntrega.Length == 0 )
				{
					MessageBox.Show("Seleccione la ubicación de entrega antes de continuar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			#endregion Validacion

			int idTipoPrecioVM = 2;
			if ((int)this.cmbFormaPago.Value == 3) idTipoPrecioVM = 3;
			decimal dPrecioOriginal = (decimal)this.txtIva0.Value + (decimal)this.txtIva.Value;
			if (!Validacion.bValidaMargen(this.ultraGrid1, idTipoPrecioVM, MenuLatinium.stUsuario, (decimal)this.txtBaseImponible.Value, dPrecioOriginal, (int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, (int)this.txtNumIdCompra.Value, 6, cdsSeteoF)) return;

			this.txtNumEstado.Value = 6;			
		
			#region Graba Factura
			try
			{
				bool bSeriales = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select FS From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);
				int idDireccion = 0;
				int idCourier = 0;
				if (this.cmbDirecciones.ActiveRow != null) idDireccion = (int)this.cmbDirecciones.Value;
				if (this.cmbCourier.ActiveRow != null) idCourier = (int)this.cmbCourier.Value;
				int idVendedor = 0;
				if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;

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
				if (this.dtpHoraEntrega.Value != System.DBNull.Value)
				{
					bHorEnt = true;
					dtHoraEntrega = (DateTime) this.dtpHoraEntrega.Value;
				}
				if (bCrediPoint) bFecPrimPag = true;
				#endregion Fechas
					
				this.txtNumIdCompra.Value = 
					Facturacion.GrabaFacturacion((int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtIdSucursal.Value, 
					(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
					(int)this.txtNumEstado.Value, bConsignacion, (int)this.txtNumContadoCredito.Value, (int)this.cmbFormaPago.Value,
					(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, 
					idVendedor, this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
					(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
					(DateTime)this.dtFecha.Value, this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad, 
					(bool)this.chkEntregado.Checked, dtFechaEnt, dtHoraEntrega, idDireccion, this.txtNotas.Text.ToString(),
					bCrediPoint, dTotalCuotas, NumCuotas, Convert.ToDecimal(this.txtCuotaDeEntrada.Value), dtFechPrimPago,
					(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
					bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bAnulado, true, (int)this.cmbBodega1.Value, 0, "", 0.00m,
					dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, false, dValorCuota, (int)this.txtIdEF.Value, 
					(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, "", bVeriricador, 
					(int)this.optTipoVenta.EditValue, this.txtNumTransaccion.Text.ToString(), this.txtDireccionEntrega.Text.ToString(), idCourier, this.txtNumGuia.Text.ToString(), Convert.ToDecimal(this.txtValorEntrega.Value),
					(int)this.txtNumIdproyectoZoom.Value, this.txtClienteNombre.Text.ToString(), 0.00m, bNuevo, bNumeracionAutomatica, idBloqueaTransacciones, this.ultraGrid1, bSeriales, idTipoOrigen, this.grdSerial, bDescuentoActivo);										
							
				if (idMail > 0)
				{
					string sSQLLM = string.Format("Update listacorreosmarketing Set idCompra = {0} Where idCorreosMarketing = {1}", (int)this.txtNumIdCompra.Value, idMail);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLLM);
				}

				string sSQLLMenterga = string.Format("Update compra Set EntregadoLocal = {0}, CodigoEntrega = '{1}' Where idcompra = {2}", Convert.ToInt32( this.chkEntregaLocal.Checked) , CodigoEntrega,(int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLMenterga);

				Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, (int)this.txtNumIdCompra.Value, cdsSeteoF);

				#region FamilyFriends
				FamilyFriendsCompraVenta(FamilyFriendsService.EstadoFFCompraTemporal);
				#endregion FamilyFriends

				//Actualizar encuesta cuando se es temporal
				if((int)this.txtNumIdCompra.Value > 0)
				{
					if(iEncuesta >0)
					{
						string squery = String.Format("Update Compra Set Encuesta = {0}, OtrosEncuesta = '{1}' where idCompra = {2}", iEncuesta, sOtrosEncuesta, (int)this.txtNumIdCompra.Value);
						Funcion.EjecutaSQL(cdsSeteoF,squery);
					}
				}

				idCuponWeb = 0;
						
				#region Audita Descuento
				if (bDescuentoActivo)
				{
					string sMensaje = "";
					
					int iContador = 1;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						if (iContador > 1) sMensaje = " - ";

						sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

						iContador++;
					}

					string sSQLCD = string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'FACTURA', {1}, 2, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCD);						
				}
				#endregion Audita Descuento
					
				#region Controles
				this.cmbFormaPago.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbBodega1.Enabled = false;
				this.cmbVendedor.Enabled = false;
				this.cmbTipoRuc.Enabled = false;

				this.txtNumero.Enabled = false;
				this.txtRuc.Enabled = false;
				this.txtRuc1.Enabled = false;
				this.dtFecha.Enabled = false;
				this.cmbFechaPrimerPago.Enabled = false;
				this.cmbNumContadoCredito.Enabled = false;

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
				this.btnValidarMail.Enabled = false;
													
				HabilitaTab(false);
				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
				FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);

				bNuevo = false;
				bEdicion = false;
				bDescuentoActivo = false;
				idBloqueaTransacciones = 0;
				dClavePorcDscto = 0.00m;
				bCuponWeb = false;
				this.Consultar((int)this.txtNumIdCompra.Value);
				SettearFamilyFriends();
				#endregion Controles				

				if (bServicioTecnico) this.DialogResult = DialogResult.Yes;
			}
			catch (Exception ex)
			{				
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
			}
			#endregion Graba Factura
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
			if (this.chkEntregado.Checked)
			{
				if ((int)this.txtNumIdCliente.Value == 0)
				{
					MessageBox.Show("Ingrese un Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.chkEntregado.Checked	= false;
					this.txtRuc.Focus();
					return;
				}

				if ((int)this.optTipoVenta.EditValue == 1)
				{
					if((int)this.cmbTipoRuc.Value==11)
					{
					}
					else 
					{
						if (this.cmbDirecciones.Rows.Count == 0)
						{
							MessageBox.Show("El Cliente No Tiene Direcciones para la Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.chkEntregado.Checked	= false;
							this.txtRuc.Focus();
							return;
						}
					}
				}
				this.dtFechaEntrega.Enabled = true;
				this.dtpHoraEntrega.Enabled = true;
				this.cmbDirecciones.Enabled = true;
				this.cmbCourier.Enabled = true;		
				this.txtNumTransaccion.Enabled=true;
				this.txtresponsablecurier.Enabled=true;	
				this.txtDireccionEntrega.Enabled = true;
				this.txtNumGuia.Enabled = true;
				this.txtValorEntrega.Enabled = true;
				this.dtphoracurier.Enabled=true;
				this.dtFechaEntregaCuriermbo1.Enabled=true;
	
				this.dtFechaEntrega.Value = (DateTime)this.dtFecha.Value;
						
				this.dtFechaEntrega.Focus();
				this.chkEntregaLocal.Enabled = false;
				this.btnEntrega.Visible = true;
				if(bEdicion || bNuevo)this.btnEntrega.Enabled = true;
				this.chkEntregaLocal.Checked = false;
				#region gps
				// Hora actual en UTC
				DateTime ahora = DateTime.UtcNow;

				// Fecha base del Unix epoch
				DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0);

				// Calcula los segundos desde el epoch
				long unixTime = (long)(ahora - epoch).TotalSeconds;

				// Guarda en variable string
				if(CodigoEntrega.Length == 0) CodigoEntrega = unixTime.ToString();
				#endregion gps

			}
			else
			{
				this.dtFechaEntrega.Value = System.DBNull.Value;
				this.dtpHoraEntrega.Value = System.DBNull.Value;
				this.cmbDirecciones.Value = System.DBNull.Value;
				this.cmbCourier.Value = System.DBNull.Value;
				this.txtDireccionEntrega.Text = "";
				this.txtNumGuia.Text = "";
				this.txtValorEntrega.Value = 0.00m;

				this.dtFechaEntrega.Enabled = false;
				this.dtpHoraEntrega.Enabled = false;
				//this.cmbDirecciones.Enabled = false;
				this.cmbCourier.Enabled = false;							
				this.txtDireccionEntrega.Enabled = false;
				this.txtNumGuia.Enabled = false;
				this.txtValorEntrega.Enabled = false;
				this.chkEntregaLocal.Enabled = true;
				this.btnEntrega.Visible = false;
				this.btnEntrega.Enabled = false;
			}
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			//			if (bEdicion)
			//				if (!Validacion.vbNumeracion(bNumeracionAutomatica, (int)this.txtNumIdCompra.Value, 1, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) e.Cancel = true;			
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) 
			{
				if (this.txtRuc.Enabled) this.txtRuc.Focus();
				else 
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
			}
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
				bool bConsignacion = false;
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
				if (ultraGrid1.ActiveRow == null)
				{
					MessageBox.Show("Seleccione el Artículo", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}

				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Focus();
					return;
				}
				#endregion Valida Selección 

				int IdArticulo = (int) this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
				int IPosicion = (int) this.ultraGrid1.ActiveRow.Cells["Posicion"].Value; 
				
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
					string sSerial = dr.Cells["Serial"].Value.ToString();
					if ((string)this.txtSerial.Text == sSerial)
					{
						MessageBox.Show("Ya registro esta serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
								
				#region Valida si la Serial esta en Orden de Compra - En Transito - Pedidos - Reservas - Consignaciones - Ventas 
				int iCont = 0;
				string sMens = "";
				string sSQLValSerEsp = string.Format("Exec ValidaSerialOrdCmpEnTransito '{0}', {1}, {2}, 1, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega1.Value, idCompraOrigen);
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

				#region Valida Serie Bloqueada
				string sSQLVSEB = string.Format("Exec ValidaSerialBloqueada '{0}'", this.txtSerial.Text.ToString());
				string sMensajeVSEB = Funcion.sEscalarSQL(cdsSeteoF, sSQLVSEB);
				
				if (sMensajeVSEB.Length > 0)
				{
					MessageBox.Show(sMensajeVSEB, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
				#endregion Valida Serie Bloqueada 

				#region Carga Grid de Seriales
				string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega1.Value, idCompraOrigen);
				int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
				if (iCtaSerial == 1)
				{	
					IdSerial = 0;
					IdDetalleSerial = 0;

					string sSQLIdSerial = string.Format("Exec BuscaIdSerial '{0}', {1}, {2}, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega1.Value, idCompraOrigen);
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						IdSerial = drIdSerial.GetInt32(0);
						IdDetalleSerial = drIdSerial.GetInt32(1);
						bConsignacion = drIdSerial.GetBoolean(2);
					}
					drIdSerial.Close();
					
					try 
					{
						#region serial 
						if (!ModSerial) 
						{		
							//UltraGridRow newRow = this.grdSerial.Rows.Band.AddNew();										
							this.grdSerial.Rows.Band.AddNew();
							this.grdSerial.ActiveRow.Cells["idSerial"].Value = IdSerial;
							this.grdSerial.ActiveRow.Cells["idDetalleSerial"].Value = 0;
							this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
							this.grdSerial.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
							this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
							this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
							this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
							this.grdSerial.ActiveRow.Cells["Consignacion"].Value = bConsignacion;
						}
						else
						{
							this.grdSerial.ActiveRow.Cells["idSerial"].Value = IdSerial;						
							this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;						
							this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
							this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
							this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
							this.grdSerial.ActiveRow.Cells["Consignacion"].Value = bConsignacion;
						}
						#endregion serial 
					}
					catch(System.Exception ex)
					{
						MessageBox.Show(ex.Message, "Point Technology");
					}
						
				}
				else
				{
					MessageBox.Show("La serial no existe, Esta en otro Local o no pertenece al artículo seleccionado,\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Focus();
					return;
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
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

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
				if (idTipoOrigen == 3 || idTipoOrigen == 11 || idTipoOrigen == 13)
				{
					MessageBox.Show("No puede Eliminar seriales de Pedidos-Reservas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					e.Cancel = true;
					return;
				}

				FuncionesInventario.EliminaFilasGridsSeriales(sender, e, 1, (int)this.txtNumIdCompra.Value, 0, cdsSeteoF);							
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
			if (bEdicion)
			{			
				if (e.Cell.Column.ToString() == "Codigo")
				{
					if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0) e.Cancel = true;

					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
					{			
						if (!bModificaColor && (int)e.Cell.Row.Cells["idArticulo"].Value > 0)
						{
							MessageBox.Show("Para reemplazar un articulo, elimine el articulo actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
						}
					}

					if (!Validacion.bValidaCambioArticuloVenta(e, bModificaColor, false))
					{
						e.Cancel = true;
						return;						
					}
				}				
				
				if (e.Cell.Column.ToString() == "Precio")
				{
					//if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0) e.Cancel = true;

					if (e.Cell.Row.Cells["Tipo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 5 || (int)e.Cell.Row.Cells["Tipo"].Value == 8 || (int)e.Cell.Row.Cells["Tipo"].Value == 9 || (int)e.Cell.Row.Cells["Tipo"].Value == 10 || (int)e.Cell.Row.Cells["Tipo"].Value == 11) e.Cancel = true;

					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["idArticulo"].Value == 8335) e.Cancel = true;
				}

				if (e.Cell.Column.ToString() == "Cantidad")
				{
					if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0) e.Cancel = true;
				
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["idArticulo"].Value == 8335) e.Cancel = true;
				}

				if (e.Cell.Column.ToString() == "Impuesto")
				{
					if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0) e.Cancel = true;

					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["idArticulo"].Value == 8335) e.Cancel = true;
				}
			
				if (e.Cell.Column.ToString() == "DescuentoPorc") 
				{
					if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0 && idTipoOrigen != 5) e.Cancel = true;				

					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if (e.Cell.Row.Cells["Tipo"].Value != System.DBNull.Value)
							if ((int)e.Cell.Row.Cells["Tipo"].Value == 5 || (int)e.Cell.Row.Cells["Tipo"].Value == 8 || (int)e.Cell.Row.Cells["Tipo"].Value == 9 || (int)e.Cell.Row.Cells["Tipo"].Value == 10) e.Cancel = true;
					//if ((int)e.Cell.Row.Cells["idArticulo"].Value == 8335) e.Cancel = true;
				
					Validacion.ModificaDescuentos(e, false, false);
				}
			}
		}

		bool bModificaColor = false;

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			Infragistics.Win.UltraWinGrid.UltraGridRow row = this.ultraGrid1.ActiveRow;

			if (row != null && row.Cells["idArticulo"] != null && row.Cells["idArticulo"].Value != DBNull.Value)
			{
				int idArticulo = Convert.ToInt32(row.Cells["idArticulo"].Value);
				
				//if (idArticulo == 12920 || idArticulo == 12921) 
				if (idArticulo == 12929 || idArticulo == 12930) 	
				{
					FuncionesInventario.AsistenciaSaludDental(this.ultraGrid1, Convert.ToDateTime(this.dtFecha.Value), this.ultraGrid1.Rows.Count - 1, cdsSeteoF, (int)this.cmbFormaPago.Value, bEmpleado);
				}
			}

			if (bEdicion)
			{
				if ((int)this.cmbFormaPago.Value != 9)
				{
					if (bEdicion)
					{
						if (!FuncionesInventario.CambioColorSerialesIngresadas(this.ultraGrid1, this.grdSerial)) return;

						bModificaColor = true;
						Facturacion.CambiaColorProducto(this.ultraGrid1, 1, (int)this.cmbBodega.Value, cdsSeteoF);
						bModificaColor = false;				
					}
				}
			}
		}

		private void dtFechaEntrega_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbDirecciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnClaveDescuento_Click(object sender, System.EventArgs e)
		{
			using (frmVentaClaveDescuento Clave = new frmVentaClaveDescuento())
			{				
				if (DialogResult.OK == Clave.ShowDialog())
				{
					#region Audita Descuento
					bDescuentoActivo = true;

					string sMensaje = "";
					
					int iContador = 1;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						if (iContador > 1) sMensaje = " - ";

						sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

						iContador++;
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'FACTURA', {1}, 1, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje));
					dClavePorcDscto = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaPorcentajeClaveDescuento {0}", MenuLatinium.IdUsuario));
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set UsuarioDescuento = '{0}', DescuentoActivo = 1 Where idCompra = {1}", MenuLatinium.stUsuario, (int)this.txtNumIdCompra.Value));
					#endregion Audita Descuento
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
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (idTipoOrigen == 3) return;

				if (this.dtFecha.Value != System.DBNull.Value)
				{														
					if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					if (this.cmbBodega.ActiveRow != null && idTipoOrigen != 3) this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
					
					this.dtFechaRetencion.Value = (DateTime)this.dtFecha.Value;					
				}
			}

			if (this.dtFecha.Value != System.DBNull.Value)
			{
				this.dtFechaRetencion.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
				this.dtFechaRetencion.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(5);	
				this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today;
				this.dtFechaEntrega.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(3);	
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "VENTA", cdsSeteoF)) e.Cancel = true;			
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				if (bEdicion && idTipoOrigen == 3) return;
				
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtFecha.Value, false, cdsSeteoF);

				this.dtFechaEntrega.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;				
			}
		}

		private void cmbCaducidad_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void mnuImprimirFactura_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);//-2
			
		}

		private void mnuDocsCrediPoint_Click(object sender, System.EventArgs e)
		{
			decimal dValor = (decimal)this.txtTotal.Value - (decimal)this.txtCuotaDeEntrada.Value;
			string Fecha = dtFechPrimPago.ToString();//.ToString("yyyyMMdd");			
			DocsPrefactura Dctos = new DocsPrefactura ((int)this.txtNumIdCompra.Value, dValor, NumCuotas, (int)this.txtNumIdCliente.Value, Fecha, (DateTime)this.dtFecha.Value, 1, this.txtRuc.Text.ToString());
			Dctos.ShowDialog();		
		}

		private void mnuImprimirGarantiaExtendida_Click(object sender, System.EventArgs e)
		{
			this.ImpresionGarantiaExtendida();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void utcTransaccion_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void optTipoVenta_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				
				this.txtNumTransaccion.Enabled  = true; 
				this.txtNumGuia.Enabled = true;
				this.txtresponsablecurier.Enabled = true;
				this.txtValorEntrega.Enabled = true;
				//this.txtpieza.Enabled = true;
				this.dtFechaEntregaCuriermbo1.Enabled = true;
				this.dtphoracurier.Enabled=true;
				this.cmbCourier.Enabled = true;
				this.txtDireccionEntrega.Enabled = true;
				this.dtFechaEntrega.Enabled=true;
				this.dtpHoraEntrega.Enabled=true;

				
			}
		}

		private void txtValorEntrega_Click(object sender, System.EventArgs e)
		{
			this.txtValorEntrega.SelectAll();
		}

		private void txtValorEntrega_Enter(object sender, System.EventArgs e)
		{
			this.txtValorEntrega.SelectAll();
		}

		private void txtNumTransaccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtNumGuia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void mnuBICotizaciones_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(3);
		}

		private void mnuBIPedidos_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(13);
		}

		private void txtProyecto_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void txtClienteNombre_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				string sSQL = string.Format("Exec VentasNumeroDeFacturas '{0}'", this.txtClienteNombre.Text.ToString());
				iNumCompras = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			}
		}

		private void txtProyecto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void txtClienteNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void txtProyecto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtProyecto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtClienteNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtfechacurier_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumTransaccion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnCliente2_Click(object sender, System.EventArgs e)
		{
			this.txtRuc1.Text=FuncionesProcedimientos.AbreFacturacionClientes1(false, 1, (int)this.txtNumIdCliente1.Value);
			CargaCliente();			 
			this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtNumIdCliente1.Value, cdsSeteoF);
		}

		private void txtNumIdCliente_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumIdCliente1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtClienteNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label61_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnImprimeED_Click(object sender, System.EventArgs e)
		{
			if ((bool) this.chkEntregado.Checked)
			{        
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idCompra";
				ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				Reporte miRep = new Reporte("ReporteEntregaDomicilio.Rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Recepción Mercaderia de Retiro");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();			
			}
		}

		int idMail = 0;

		private void btnValidarMail_Click(object sender, System.EventArgs e)
		{
			idMail = Facturacion.iPromocionMail(cdsSeteoF);
		}

		private void mnuImprimirDYD_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				if ((int)dr.Cells["idArticulo"].Value == 8683) iCont++;
  	
			if (iCont > 0)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = (int)this.txtNumIdCompra.Value; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte("DatosSeguroDesempleo.rpt", "");//ReporteProteccionDYD.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Proteccion D&D");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();				
			}
		}

		private void mnuImprimirRIDE_Click(object sender, System.EventArgs e)
		{
			if (bAnulado)	
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) != 10)	
			{
				MessageBox.Show("El documento no esta Contabilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFe_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) != 2)	
			//			{
			//				MessageBox.Show("El documento no ha sido autorizado por el S.R.I.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				return;
			//			}

			#region Nombre Reporte
			string sNomReporte = "";
			sNomReporte = sRideFactura;

			#region EsMoto
			if(MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 3)
			{
				// SOLO SI ES MOTO
				int idCompraMoto = 0;

				if(this.txtNumIdCompra.Value != DBNull.Value && this.txtNumIdCompra.Value != null)
				{
					idCompraMoto = Convert.ToInt32(this.txtNumIdCompra.Value);
				}

				if(bEsMoto(idCompraMoto))
				{
					sNomReporte = "Doc_FelecMotos.rpt";
				}
			}
			#endregion

			if(MenuLatinium.stIdDB == 26) sNomReporte = "RIDEFacturaZoommatika.rpt";
			#endregion Nombre Reporte

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
			discreteValIdDetCompra.Value = (int)this.txtNumIdCompra.Value; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			Reporte Reporte = new Reporte(sNomReporte, "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("RIDE");
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();				
		}

		#region EsMoto
		private bool bEsMoto(int idCompra)
		{
			try
			{
				int iEsMoto = Funcion.iEscalarSQL(
					cdsSeteoF,
					string.Format("EXEC CompraEsMoto {0}", idCompra),false
					);

				return (iEsMoto == 1);
			}
			catch
			{
				return false;
			}
		}

		#endregion

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFe_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) == 2)
			{
				this.timer1.Enabled = false;

				MessageBox.Show("Comprobante Autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramFieldIdDetCompra = new ParameterField ();
				ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
				paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
				discreteValIdDetCompra.Value = (int)this.txtNumIdCompra.Value; 
				paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
				paramFields.Add (paramFieldIdDetCompra);
				
				Reporte Reporte = new Reporte(sRideFactura, "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("RIDE");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();				

				this.Consultar((int)this.txtNumIdCompra.Value);
			}			
		}

		private void mnuImprimir_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void mnuDocsCrediPointBlanco_Click(object sender, System.EventArgs e)
		{
			decimal dValor = (decimal)this.txtTotal.Value - (decimal)this.txtCuotaDeEntrada.Value;
			string Fecha = dtFechPrimPago.ToString();//.ToString("yyyyMMdd");			
			frmDocsPreFacturaEnBlanco DctosBlanco = new frmDocsPreFacturaEnBlanco ((int)this.txtNumIdCompra.Value, dValor, NumCuotas, (int)this.txtNumIdCliente.Value, Fecha, (DateTime)this.dtFecha.Value, (int)this.txtIdEF.Value, this.txtRuc.Text.ToString());
			DctosBlanco.ShowDialog();
		}

		private void mnuImprimirCP_Click(object sender, System.EventArgs e)
		{
			this.ImpresionCompraProtegida();
		}

		private void mnuImprimirPR_Click(object sender, System.EventArgs e)
		{
			this.ImpresionPointRespaldo();
		}

		private void mnuImprimirASD_Click(object sender, System.EventArgs e)
		{
			this.ImpresionAsistenciaSaludDental();
		}


		private void mnuImprimirGEP_Click(object sender, System.EventArgs e)
		{
			this.ImpresionGarantiaExtendidaPlus();
		}

		private void lblBodega_Click(object sender, System.EventArgs e)
		{
		
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(11);
		}

		private void mnuBILaboratorio_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			PedidoReserva(69);
		}

		private void ultraDataSource2_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			#region Imprime Celulares 
			if ((int)this.cmbFormaPago.Value == 9)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idCompra";
				ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				Reporte miRep = new Reporte("ContratoCompraVentaCelular.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Recepción Mercaderia de Retiro");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();	
			}
			#endregion Imprime Celulares 
		}

		private void mnuImprimirNuovo_Click(object sender, System.EventArgs e)
		{
			int icolateral = 0;
			icolateral = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select (SELECT CONVERT (INT, ISNULL (bColateral,0)) FROM DBO.Compra WHERE idCompra = c.idOrigen) from dbo.Compra c where idCompra = {0}",
				(int) this.txtNumIdCompra.Value));
	
			if (icolateral == 0) 
			{
				MessageBox.Show("Reporte para clientes con colateral.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\ContratoCompraVentaCelularNuovo.rpt";

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldId = new ParameterField ();
			ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
			paramFieldId.ParameterFieldName = "@idCompra";
			ValIdId.Value = int.Parse(this.txtNumIdCompra.Value.ToString());
			paramFieldId.CurrentValues.Add (ValIdId);
			paramFields.Add (paramFieldId);

			Reporte miRep = new Reporte("ContratoCompraVentaCelularNuovo.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("contrato Colateral");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();	

		}

		private void Fundas()
		{
			string sIdArticulos = "";
			int iFila = 0;
			#region Articulos en el Grid
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (sIdArticulos.Length == 0) sIdArticulos = dr.Cells["idArticulo"].Value.ToString();
					else if (sIdArticulos.Length > 0) sIdArticulos = sIdArticulos + ", " + dr.Cells["idArticulo"].Value.ToString();										
				}
			}
			#endregion Articulos en el Grid
			if (sIdArticulos.Length > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaArticulosFundas '{0}'", sIdArticulos)) > 0)
				{
					using (frmFundasProductos GE = new frmFundasProductos(this.ultraGrid1))
					{
						if (DialogResult.OK == GE.ShowDialog())
						{
							if (GE.ultraGrid1.Rows.Count > 0)
							{
								bFundas = true;

								#region Carga Productos
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in GE.ultraGrid1.Rows.All)
								{
									if ((bool)dr.Cells["Sel"].Value)
									{
										int iArticulo =(int)dr.Cells["idArticulo"].Value ;

										#region Carga Fundas	
										decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

										this.ultraGrid1.DisplayLayout.Bands[0].AddNew();	
										iFila++;
										#region Asignacion Varibles				
										string sSQLInf = string.Format("Select a.idArticulo,a.Codigo,a.Articulo, Precio3 From Articulo a Where idArticulo = {0}", iArticulo);
										SqlDataReader drA = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
										drA.Read();

										this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["idUnidad"].Value = 1;			
										this.ultraGrid1.ActiveRow.Cells["idCombo"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["Tipo"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drA.GetString(1);			
										this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = drA.GetString(2);
										this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = drA.GetInt32(0);
										this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value = 1;										
										this.ultraGrid1.ActiveRow.Cells["Entregado"].Value = false;
										this.ultraGrid1.ActiveRow.Cells["idArticuloGE"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["Principal"].Value = false;			
										this.ultraGrid1.ActiveRow.Cells["Bloqueado"].Value = false;
										this.ultraGrid1.ActiveRow.Cells["EstadoSerial"].Value = false;
										this.ultraGrid1.ActiveRow.Cells["Posicion"].Value = this.ultraGrid1.Rows.Count;	
										this.ultraGrid1.ActiveRow.Cells["Jornadas"].Value = 0.00m;
										this.ultraGrid1.ActiveRow.Cells["Aporte"].Value = 0.00m;
										this.ultraGrid1.ActiveRow.Cells["idSubProyecto"].Value = 0;
										this.ultraGrid1.ActiveRow.Cells["Precio"].Value = drA.GetDecimal(3);
										this.ultraGrid1.ActiveRow.Cells["DescuentoPorc"].Value = 70.00m;
										this.ultraGrid1.ActiveRow.Cells["DescuentoArt"].Value = (drA.GetDecimal(3) * 70.00m)/100.00m;
										this.ultraGrid1.ActiveRow.Cells["Impuesto"].Value = dIVA;

										drA.Close();	
										#endregion Asignacion Varibles				
										#endregion Carga Articulo Fundas
									}
								}
								bFundas=false;
								#endregion Carga Productos
							}
						}
					}
				}
			}
		}
		private void txtRuc_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbVendedor.ActiveRow != null)
			{
				UltraGridRow activeRow = cmbVendedor.ActiveRow;
				// Obtener el valor de la columna 'codigo' de la fila activa
				string codigo = activeRow.Cells["codigo"].Text;
				if ( codigo == "FREELANCE")
				{
					this.chkFreelance.Visible = true;
					this.chkFreelance.Checked = true;
				}
				else
				{
					this.chkFreelance.Visible = false;
					this.chkFreelance.Checked = false;
				}
			}
		}
		private string sUrlNube = "";
		private void btnSubir_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
						
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) return;							
			#endregion Validacion

			//string carpetaBase = @"\\192.168.1.10\Proyectos\FacturasDeUna";
			string carpetaBase = "\\\\" + MenuLatinium.DirectorioArchivo + "\\Proyectos\\FacturasDeUna";

			Funcion.SubirArchivo("DEUNA", this.txtRuc.Text.ToString(), (int)this.cmbBodega.Value, true, carpetaBase, ref sUrlNube); 
			
		}
		private bool datosInicialesCargados = false; 

		private void EjecutarProcesoPagados()
		{
			try
			{
				DataView dvFormasPago = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idFormaPago, FormaPago FROM CompraFormaPago WHERE idFormaPago = 21");

				this.cmbFormaPago.DataSource = dvFormasPago;
				this.cmbFormaPago.DisplayMember = "FormaPago";
				this.cmbFormaPago.ValueMember = "idFormaPago";

				if (dvFormasPago.Count > 0)
					this.cmbFormaPago.Value = 21;
				else
				{
					MessageBox.Show("No se pudo cargar la forma de pago 21", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				this.txtRuc.Text = CedulaP;
				this.txtComprobante.Text = NFactuiraP;
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

				int idPersonal = 0;
				int bodegaVendedor = 1;

				if ((int)this.txtBodegaPredef.Value == 0)
				{
					using (SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("SELECT TOP 1 C.Bodega, C.idPersonal FROM Com_AsignacionDeVendedores C INNER JOIN Nomina N ON N.idPersonal = C.idPersonal WHERE NIdentificacion = '{0}' order by Hasta desc", SCedulaVendedor), true))
					{
						if (dr.Read())
						{
							bodegaVendedor = dr.GetInt32(0);
							idPersonal = dr.GetInt32(1);
						}
					}
					
					this.cmbBodega.Enabled = false;
					this.cmbBodega.Value = bodegaVendedor;
					NuevoRegistro(this, EventArgs.Empty, 21);
					cmbBodega_ValueChanged(this.cmbBodega, EventArgs.Empty);
					this.txtRuc.Text = CedulaP;
				}
				else
				{
					bodegaVendedor = (int)this.txtBodegaPredef.Value;
					this.cmbBodega.Value = bodegaVendedor;
				}

				DataView dvBodegas1 = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 0", bodegaVendedor));
				this.cmbBodega1.DataSource = dvBodegas1;
				this.cmbBodega1.DisplayMember = "Nombre";
				this.cmbBodega1.ValueMember = "Bodega";

				if (dvBodegas1.Count > 0)
				{
					this.cmbBodega1.Value = dvBodegas1[0]["Bodega"];
				}
			
				if (idPersonal > 0)
				{
					this.idPersonalGlobal = idPersonal; 
					this.BeginInvoke(new MethodInvoker(this.RestaurarVendedor));
				}

				this.txtRuc.Enabled = false;
				this.txtComprobante.Enabled = false;

				if (!datosInicialesCargados)
				{
					this.BeginInvoke(new MethodInvoker(this.EjecutarValidacionRuc));
					datosInicialesCargados = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error en el proceso automático: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void DeferredValidated(object sender, EventArgs e)
		{
			EjecutarValidacionRuc();
		}
		private void RestaurarVendedor()
		{
			foreach (UltraGridRow row in this.cmbVendedor.Rows)
			{
				if (((int)row.Cells["idPersonal"].Value) == idPersonalGlobal)
				{
					this.cmbVendedor.Value = idPersonalGlobal;
					this.cmbVendedor.Enabled = false;
					break;
				}
			}
		}

		private void EjecutarValidacionRuc()
		{
			this.txtRuc_Validated(this.txtRuc, EventArgs.Empty);
		}
		private void EjecutarProcesoRebates()
		{
			FuncionesInventario._modoRebates = true;
 
			if (this.btnNuevo != null && this.btnNuevo.Enabled)
				this.btnNuevo.PerformClick();
 
			if (this.cmbBodega != null) this.cmbBodega.Value  = 1;
			if (this.cmbBodega1 != null) this.cmbBodega1.Value = 1;
			if (this.cmbTipoRuc != null) this.cmbTipoRuc.Value = 4;
			if (this.cmbVendedor != null) this.cmbVendedor.Value = 4044;

			this.txtNumIdCliente.Value = idClienteR;
			this.txtNombre.Value = clienteN;
			if (this.txtRuc != null) this.txtRuc.Text  = (this.cedula == null) ? "" : this.cedula.Trim();
			if (this.txtNotas != null) this.txtNotas.Value = "FACTURA GENERADA DESDE LIQUIDACION REBATES";

			if (this.txtNumIdCompra != null && (this.txtNumIdCompra.Value == null || this.txtNumIdCompra.Value == System.DBNull.Value))
				this.txtNumIdCompra.Value = 0;
 
			this.ultraGrid1.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
 
			Infragistics.Win.UltraWinGrid.UltraGridBand band = this.ultraGrid1.DisplayLayout.Bands[0];
			Infragistics.Win.UltraWinGrid.UltraGridRow row = band.AddNew();

			if (row.Cells.Exists("idArticulo")) row.Cells["idArticulo"].Value = System.DBNull.Value;
			if (row.Cells.Exists("Codigo")) row.Cells["Codigo"].Value = "";
			if (row.Cells.Exists("Articulo")) row.Cells["Articulo"].Value = "";

			bEdicion = true;
			bActualiza = true;
			bCargaPromocionFP = false;
			bFundas = false;
 
			if (this.cmbFormaPago != null && this.cmbFormaPago.Value == null)
				this.cmbFormaPago.Value = 4;

			int idArtObjetivo = 8684; 
			string sCodigoArt = Funcion.sEscalarSQL(cdsSeteoF, "Select Codigo From Articulo Where idArticulo = " + idArtObjetivo );
			if (sCodigoArt == null) sCodigoArt = "";
			sCodigoArt = sCodigoArt.Trim();

			decimal IvaVigente = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
 
			this.ultraGrid1.ActiveRow  = row;
			this.ultraGrid1.ActiveCell = row.Cells["Codigo"];

			this.ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode);
			row.Cells["Codigo"].Value = sCodigoArt + " ";  
			this.ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);

			this.ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode);
			row.Cells["Codigo"].Value = sCodigoArt;     
			this.ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
			 
			Infragistics.Win.UltraWinGrid.UltraGridRow r = this.ultraGrid1.ActiveRow;
			bActualiza = false;

			if (r != null)
			{
				if (r.Cells.Exists("Cantidad") && (r.Cells["Cantidad"].Value == null || r.Cells["Cantidad"].Value == System.DBNull.Value))
					r.Cells["Cantidad"].Value = 1;

				if (r.Cells.Exists("DescuentoArt")  && (r.Cells["DescuentoArt"].Value  == null || r.Cells["DescuentoArt"].Value  == System.DBNull.Value))
					r.Cells["DescuentoArt"].Value = 0m;

				if (r.Cells.Exists("DescuentoPorc") && (r.Cells["DescuentoPorc"].Value == null || r.Cells["DescuentoPorc"].Value == System.DBNull.Value))
					r.Cells["DescuentoPorc"].Value = 0m;

				if (r.Cells.Exists("Precio"))   r.Cells["Precio"].Value   = valorR;   
				if (r.Cells.Exists("Impuesto")) r.Cells["Impuesto"].Value = IvaVigente;  
				r.Update();
			}

			this.Total();
			if (this.btnGuardar != null && this.btnGuardar.Enabled) this.btnGuardar.PerformClick();
		}

		private void pgEntrega_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void chkEntregaLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkEntregaLocal.Checked)
			{
				this.chkEntregado.Checked = false;
				this.chkEntregado.Enabled = false;
			}
			else
			{
				this.chkEntregado.Enabled = true;
			}
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
		
		}

		private void btnEntrega_Click(object sender, System.EventArgs e)
		{
			// Obtiene el valor seleccionado de la bodega
			int bodegaGps = (int)this.cmbBodega.Value;

			// Obtiene la cédula del textbox
			string cedulaGps = (string)this.txtRuc.Text;

			// URL base del servidor
			string sUrlNube = "https://credipoint.appservices.com.ec/mapScreenLatinium/";

			// Variable que debes tener definida antes (por ejemplo, string CodigoEntrega = "123456";)
			string linkUrl = sUrlNube + bodegaGps + "/" + cedulaGps + "/geolocalizacion/" + CodigoEntrega;

			// Valida que la URL no esté vacía antes de abrirla
			if (sUrlNube.Length > 10)
			{
				System.Diagnostics.Process.Start("chrome.exe", linkUrl);
			}
		}

		private void txtCuotaDeEntrada_ValidationError(object sender, Infragistics.Win.UltraWinEditors.ValidationErrorEventArgs e)
		{
		
		}


		private void cmbNumContadoCredito_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbFormaPago.Value == 35)
				{				
					NumCuotas = 0;

					if (this.cmbNumContadoCredito.ActiveRow != null)						
						NumCuotas = (int) this.cmbNumContadoCredito.Value;					
				}
			}
		}

		private void cmbBodega_AfterCloseUp(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null && bNuevo)
			{	
				#region Integración Yupana (Forma de Pago 35)
				if (bNuevo && this.cmbFormaPago.Value != null && this.cmbFormaPago.Value != DBNull.Value && (int)this.cmbFormaPago.Value == 35)
				{
					bool bRegistroSeleccionado = false;
					using (frmBuscaCre_SolicitudYupana frmYupana = new frmBuscaCre_SolicitudYupana((int)this.cmbBodega.Value))
					{
						if (frmYupana.ShowDialog() == DialogResult.OK && frmYupana.ultraGrid1.ActiveRow != null)
						{
							if (frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value != null && frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value != DBNull.Value)
							{
								this.cmbTipoRuc.Value =  Convert.ToInt32( frmYupana.ultraGrid1.ActiveRow.Cells["idTipoRuc"].Value);
								this.txtRuc.Text = frmYupana.ultraGrid1.ActiveRow.Cells["Cedula"].Value.ToString();
								this.txtRuc_Validated(this.txtRuc, System.EventArgs.Empty);
								idCre_SolicitudWeb = Convert.ToInt32( frmYupana.ultraGrid1.ActiveRow.Cells["idCre_SolicitudWeb"].Value);
								this.cmbNumContadoCredito.Enabled = true;
								this.cmbFechaPrimerPago.Enabled = true;
								this.txtComprobante.Text = Convert.ToString( frmYupana.ultraGrid1.ActiveRow.Cells["NumeroSolicitud"].Value);
								this.lblYupana.Text ="YUPANA";
								this.cmbVendedor.Value = Convert.ToInt32( frmYupana.ultraGrid1.ActiveRow.Cells["idVendedor"].Value);
                 bCrediPoint = true;
								bRegistroSeleccionado = true;
							}
						}
					}

					if (!bRegistroSeleccionado)
					{
						MessageBox.Show("En la bodega seleccionada no existe registro de solicitud Yupana", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.cmbBodega.Value = System.DBNull.Value;
					}
				}
				#endregion
			}
		}


		private void cmbFechaPrimerPago_ValueChanged(object sender, EventArgs e)
		{
			if (this.cmbFechaPrimerPago.ActiveRow != null && bNuevo &&
				this.cmbFechaPrimerPago.Value != null &&
				this.cmbFechaPrimerPago.Value != DBNull.Value)
			{
				dtFechPrimPago = Convert.ToDateTime(this.cmbFechaPrimerPago.Value);
			}
			

		}
	}
}

