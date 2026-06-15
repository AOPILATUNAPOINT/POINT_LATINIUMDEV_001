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
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDevolucionVenta.
	/// </summary>
	public class frmDevolucionVenta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFiscal;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		public DevExpress.XtraEditors.SimpleButton btnPedidoIngreso;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdproyectoZoom;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnMotivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private System.Windows.Forms.Label lblIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnAutorizar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdFormaPago;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbCaducidad;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.GroupBox gbBotones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		private System.Windows.Forms.Button btnFe;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;

		[DllImport("user32.dll")]
		static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

		private System.ComponentModel.IContainer components;

		#region Variables Globales
		int idGrupoCliente = 0;
		bool bEdicion = false;
		bool bNuevo = false;
		bool bIncluyeIva = false;
		bool bNumeroAutomatico = false;
		private bool bActualiza = true;
		bool bRastreo = false;
		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;
		bool bAsientoAutomatico = false;
		private int IdSigno = 2;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		bool bImpreso = false;
		bool bAnulado = false;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		int idBloqueaTransacciones = 0;
		string sTipoOperacion = "";
		#endregion Variables Globales

		public frmDevolucionVenta()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idCompraV = 0;

		public frmDevolucionVenta(int IdCompraV)
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDevolucionVenta));
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			this.chkFiscal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.btnPedidoIngreso = new DevExpress.XtraEditors.SimpleButton();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdproyectoZoom = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnMotivo = new System.Windows.Forms.Button();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label54 = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIva = new System.Windows.Forms.Label();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.txtNumIdFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.cmbCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnFe = new System.Windows.Forms.Button();
			this.txtidProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// chkFiscal
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFiscal.Appearance = appearance1;
			this.chkFiscal.BackColor = System.Drawing.Color.Transparent;
			this.chkFiscal.CausesValidation = false;
			this.chkFiscal.Enabled = false;
			this.chkFiscal.Location = new System.Drawing.Point(840, 72);
			this.chkFiscal.Name = "chkFiscal";
			this.chkFiscal.Size = new System.Drawing.Size(48, 22);
			this.chkFiscal.TabIndex = 144;
			this.chkFiscal.Text = "Fiscal";
			this.chkFiscal.CheckedChanged += new System.EventHandler(this.chkFiscal_CheckedChanged);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(448, 75);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 162;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(176, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 161;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance2;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(296, 72);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(144, 22);
			this.txtRuc.TabIndex = 159;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(208, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(232, 21);
			this.cmbBodega.TabIndex = 163;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(496, 72);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(336, 22);
			this.txtNombre.TabIndex = 160;
			// 
			// txtNumero
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(80, 104);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 22);
			this.txtNumero.TabIndex = 158;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 107);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 165;
			this.label1.Text = "Número";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance5.BackColor = System.Drawing.Color.White;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "_DetCompra";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 58;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 58;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 58;
			ultraGridColumn9.Header.Caption = "Código";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 250;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Artículo";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 360;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 75;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "IVA";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 50;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 60;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "% Desc";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 70;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance11;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Subtotal";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.Header.VisiblePosition = 20;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 58;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 63;
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 17;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn25});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1006, 184);
			this.ultraGrid1.TabIndex = 166;
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
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
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn17.DefaultValue = 0;
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn18.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn19.DefaultValue = false;
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn21.DefaultValue = 0;
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn22});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(86, 16);
			this.label2.TabIndex = 168;
			this.label2.Text = "N. Identificación";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(448, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 169;
			this.label3.Text = "Mótivo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPedidoIngreso
			// 
			this.btnPedidoIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnPedidoIngreso.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btnPedidoIngreso.Location = new System.Drawing.Point(798, 544);
			this.btnPedidoIngreso.Name = "btnPedidoIngreso";
			this.btnPedidoIngreso.Size = new System.Drawing.Size(48, 25);
			this.btnPedidoIngreso.TabIndex = 204;
			this.btnPedidoIngreso.Text = "Pedidos";
			this.btnPedidoIngreso.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 208;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance17;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(110, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 207;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(120, 42);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 212;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance18;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 335;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(176, 40);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(264, 21);
			this.cmbVendedor.TabIndex = 211;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			// 
			// cmbMotivos
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance19;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataSource = this.ultraDataSource1;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 272;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbMotivos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Enabled = false;
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(496, 40);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(272, 21);
			this.cmbMotivos.TabIndex = 214;
			this.cmbMotivos.ValueMember = "idMotivo";
			this.cmbMotivos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbMotivos_KeyPress);
			this.cmbMotivos.Validating += new System.ComponentModel.CancelEventHandler(this.cmbMotivos_Validating);
			this.cmbMotivos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbMotivos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbTipoRuc
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance20;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 120;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(80, 73);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(120, 21);
			this.cmbTipoRuc.TabIndex = 215;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoRuc_KeyPress);
			this.cmbTipoRuc.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipoRuc_Validating);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance21;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 352);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(656, 72);
			this.txtNotas.TabIndex = 217;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 328);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(33, 16);
			this.label7.TabIndex = 218;
			this.label7.Text = "Notas";
			this.label7.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 75);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(71, 16);
			this.label8.TabIndex = 219;
			this.label8.Text = "Identificación";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(840, 107);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(60, 16);
			this.label9.TabIndex = 220;
			this.label9.Text = "Doc Origen";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtComprobante
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance22;
			this.txtComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(904, 104);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(110, 22);
			this.txtComprobante.TabIndex = 221;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(448, 10);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(83, 16);
			this.label10.TabIndex = 224;
			this.label10.Text = "Bodega Destino";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodegaDestino
			// 
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 180;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(544, 8);
			this.cmbBodegaDestino.MaxDropDownItems = 30;
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(224, 21);
			this.cmbBodegaDestino.TabIndex = 225;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodegaDestino_KeyPress);
			this.cmbBodegaDestino.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodegaDestino_Validating);
			this.cmbBodegaDestino.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaDestino_InitializeLayout);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(816, 5);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 228;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance23;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(8, 432);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(656, 40);
			this.txtObservaciones.TabIndex = 229;
			this.txtObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObservaciones_KeyPress);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 496);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 261;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(86, 496);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.TabIndex = 260;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(164, 496);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.TabIndex = 259;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(242, 496);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.TabIndex = 258;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(320, 496);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 257;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(554, 496);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.TabIndex = 256;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(632, 496);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 255;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(710, 496);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 254;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(970, 504);
			this.txtNumIdProyecto.Name = "txtNumIdProyecto";
			this.txtNumIdProyecto.PromptChar = ' ';
			this.txtNumIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdProyecto.TabIndex = 274;
			this.txtNumIdProyecto.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(954, 504);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 273;
			this.txtIdComprobante.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(938, 504);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 272;
			this.txtNumEstado.Visible = false;
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(922, 504);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 271;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(906, 504);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 270;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(890, 504);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 269;
			this.txtIdSucursal.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(874, 504);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 266;
			this.txtNumContadoCredito.Visible = false;
			// 
			// txtNumidOrigen
			// 
			this.txtNumidOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen.Enabled = false;
			this.txtNumidOrigen.Location = new System.Drawing.Point(826, 504);
			this.txtNumidOrigen.Name = "txtNumidOrigen";
			this.txtNumidOrigen.PromptChar = ' ';
			this.txtNumidOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen.TabIndex = 265;
			this.txtNumidOrigen.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(810, 504);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 264;
			this.txtNumIdCompra.Visible = false;
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(842, 504);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 263;
			this.txtNumIdCliente.Visible = false;
			// 
			// txtNumIdAsiento
			// 
			this.txtNumIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdAsiento.Enabled = false;
			this.txtNumIdAsiento.Location = new System.Drawing.Point(858, 504);
			this.txtNumIdAsiento.Name = "txtNumIdAsiento";
			this.txtNumIdAsiento.PromptChar = ' ';
			this.txtNumIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdAsiento.TabIndex = 262;
			this.txtNumIdAsiento.Visible = false;
			// 
			// txtNumIdproyectoZoom
			// 
			this.txtNumIdproyectoZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance24.ForeColor = System.Drawing.Color.Black;
			this.txtNumIdproyectoZoom.Appearance = appearance24;
			this.txtNumIdproyectoZoom.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtNumIdproyectoZoom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdproyectoZoom.Enabled = false;
			this.txtNumIdproyectoZoom.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtNumIdproyectoZoom.ForeColor = System.Drawing.Color.Black;
			this.txtNumIdproyectoZoom.Location = new System.Drawing.Point(986, 504);
			this.txtNumIdproyectoZoom.Name = "txtNumIdproyectoZoom";
			this.txtNumIdproyectoZoom.PromptChar = ' ';
			this.txtNumIdproyectoZoom.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdproyectoZoom.TabIndex = 275;
			this.txtNumIdproyectoZoom.Visible = false;
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
			// btnMotivo
			// 
			this.btnMotivo.CausesValidation = false;
			this.btnMotivo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnMotivo.Location = new System.Drawing.Point(776, 38);
			this.btnMotivo.Name = "btnMotivo";
			this.btnMotivo.Size = new System.Drawing.Size(24, 24);
			this.btnMotivo.TabIndex = 276;
			this.btnMotivo.Text = "...";
			this.btnMotivo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnMotivo.Click += new System.EventHandler(this.btnMotivo_Click);
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance25;
			this.txtIva.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtIva.ForeColor = System.Drawing.Color.Black;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(930, 352);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 283;
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.ForeColor = System.Drawing.Color.Black;
			this.label55.Location = new System.Drawing.Point(986, 328);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 17);
			this.label55.TabIndex = 278;
			this.label55.Text = "IVA";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance26;
			this.txtDescIva0Total.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDescIva0Total.ForeColor = System.Drawing.Color.Black;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(842, 376);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 285;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance27;
			this.txtTotal.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.ForeColor = System.Drawing.Color.Black;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(842, 432);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(168, 34);
			this.txtTotal.TabIndex = 287;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance28;
			this.txtIva0.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtIva0.ForeColor = System.Drawing.Color.Black;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(842, 352);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 282;
			// 
			// label54
			// 
			this.label54.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label54.AutoSize = true;
			this.label54.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label54.ForeColor = System.Drawing.Color.Black;
			this.label54.Location = new System.Drawing.Point(890, 328);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(32, 17);
			this.label54.TabIndex = 277;
			this.label54.Text = "IVA 0";
			this.label54.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.ForeColor = System.Drawing.Color.Black;
			this.lblDescuento.Location = new System.Drawing.Point(754, 379);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(57, 17);
			this.lblDescuento.TabIndex = 279;
			this.lblDescuento.Text = "Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance29;
			this.txtDescIvaTotal.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDescIvaTotal.ForeColor = System.Drawing.Color.Black;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(930, 376);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 284;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.ForeColor = System.Drawing.Color.Black;
			this.lblIva.Location = new System.Drawing.Point(890, 403);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 17);
			this.lblIva.TabIndex = 280;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance30;
			this.txtIvaTotal.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtIvaTotal.ForeColor = System.Drawing.Color.Black;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(930, 400);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 286;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.ForeColor = System.Drawing.Color.Black;
			this.lblSubTotal.Location = new System.Drawing.Point(754, 355);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(45, 17);
			this.lblSubTotal.TabIndex = 281;
			this.lblSubTotal.Text = "Subtotal";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Black;
			this.label5.Location = new System.Drawing.Point(754, 441);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(29, 17);
			this.label5.TabIndex = 288;
			this.label5.Text = "Total";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(398, 496);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.TabIndex = 289;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.Enabled = false;
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.ImageIndex = 37;
			this.btnAutorizar.ImageList = this.imageList1;
			this.btnAutorizar.Location = new System.Drawing.Point(476, 496);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.TabIndex = 290;
			this.btnAutorizar.Text = "&Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// txtNumIdFormaPago
			// 
			this.txtNumIdFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdFormaPago.Enabled = false;
			this.txtNumIdFormaPago.Location = new System.Drawing.Point(1002, 504);
			this.txtNumIdFormaPago.Name = "txtNumIdFormaPago";
			this.txtNumIdFormaPago.PromptChar = ' ';
			this.txtNumIdFormaPago.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdFormaPago.TabIndex = 291;
			this.txtNumIdFormaPago.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(816, 42);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 292;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbCaducidad
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCaducidad.Appearance = appearance31;
			this.cmbCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbCaducidad.DateButtons.Add(dateButton2);
			this.cmbCaducidad.Enabled = false;
			this.cmbCaducidad.Format = "dd/MM/yyyy";
			this.cmbCaducidad.Location = new System.Drawing.Point(722, 105);
			this.cmbCaducidad.Name = "cmbCaducidad";
			this.cmbCaducidad.NonAutoSizeHeight = 23;
			this.cmbCaducidad.Size = new System.Drawing.Size(110, 21);
			this.cmbCaducidad.SpinButtonsVisible = true;
			this.cmbCaducidad.TabIndex = 298;
			this.cmbCaducidad.Value = ((object)(resources.GetObject("cmbCaducidad.Value")));
			this.cmbCaducidad.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.cmbCaducidad_BeforeDropDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Enabled = false;
			this.label11.Location = new System.Drawing.Point(672, 107);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(41, 16);
			this.label11.TabIndex = 297;
			this.label11.Text = "Caduca";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSerie
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance32;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Enabled = false;
			this.txtSerie.Location = new System.Drawing.Point(248, 104);
			this.txtSerie.MaxLength = 10;
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(56, 22);
			this.txtSerie.TabIndex = 296;
			this.txtSerie.ValueChanged += new System.EventHandler(this.txtSerie_ValueChanged);
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(208, 107);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(33, 16);
			this.label33.TabIndex = 295;
			this.label33.Text = "Serie ";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAutFactura
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance33;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(384, 104);
			this.txtAutFactura.MaxLength = 10;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(280, 22);
			this.txtAutFactura.TabIndex = 294;
			this.txtAutFactura.ValueChanged += new System.EventHandler(this.txtAutFactura_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Enabled = false;
			this.label31.Location = new System.Drawing.Point(312, 107);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 293;
			this.label31.Text = "Autorización";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.gbBotones.Location = new System.Drawing.Point(0, 480);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1032, 8);
			this.gbBotones.TabIndex = 625;
			this.gbBotones.TabStop = false;
			// 
			// chkConsignacion
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance34;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.CausesValidation = false;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(896, 72);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 626;
			this.chkConsignacion.Text = "Consignación";
			// 
			// btnFe
			// 
			this.btnFe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFe.CausesValidation = false;
			this.btnFe.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFe.Location = new System.Drawing.Point(792, 496);
			this.btnFe.Name = "btnFe";
			this.btnFe.TabIndex = 627;
			this.btnFe.Text = "&FE";
			this.btnFe.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFe.Visible = false;
			this.btnFe.Click += new System.EventHandler(this.btnFe_Click);
			// 
			// txtidProyecto
			// 
			this.txtidProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidProyecto.Enabled = false;
			this.txtidProyecto.Location = new System.Drawing.Point(1008, 504);
			this.txtidProyecto.Name = "txtidProyecto";
			this.txtidProyecto.PromptChar = ' ';
			this.txtidProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtidProyecto.TabIndex = 628;
			this.txtidProyecto.Visible = false;
			// 
			// frmDevolucionVenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1022, 526);
			this.Controls.Add(this.txtidProyecto);
			this.Controls.Add(this.btnFe);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.cmbCaducidad);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label54);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtSerie);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtNumIdFormaPago);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.btnMotivo);
			this.Controls.Add(this.txtNumIdproyectoZoom);
			this.Controls.Add(this.txtNumIdProyecto);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumidOrigen);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtNumIdAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.cmbMotivos);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnPedidoIngreso);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.chkFiscal);
			this.KeyPreview = true;
			this.Name = "frmDevolucionVenta";
			this.Text = "Nota de Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDevolucionVenta_KeyDown);
			this.Resize += new System.EventHandler(this.frmDevolucionVenta_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDevolucionVenta_Closing);
			this.Load += new System.EventHandler(this.frmDevolucionVenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmDevolucionVenta_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}
						
		private void UnloadMe()
		{
			this.Close();
		}
		
		private int WM_CLOSE = 0x0010;
		DateTime dtFechaAFin = DateTime.Today;
		DateTime dtFechaAIni = DateTime.Today;
		decimal dPorcIva = 0.00m;	

		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}
				
		private void frmDevolucionVenta_Load(object sender, System.EventArgs e)
		{						
			if (!Funcion.Permiso("59", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Notas de Credito Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}	

			miAcceso = new Acceso(cdsSeteoF, "0725");
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;		

			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];			
						
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
      
			#region Combos
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec PersonalVendedor 5");						
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 Order By Nombre");
			this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Inventario = 1 Order By Nombre");
			this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 Order By Motivo");
			cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");			
			#endregion Combos
			
			this.txtIdTipoFactura.Value = 5;
			this.txtIdComprobante.Value = 4;
			this.txtNumContadoCredito.Value = 2;

			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;
			
			bNumeroAutomatico = Funcion.bEscalarSQL(cdsSeteoF, ("Select Automatico From CompraNumero Where idTipoFactura = 5"), true);
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Now;

			if (idCompraV > 0) this.Consultar(idCompraV);
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 5);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}		

		#region Ultragrid
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
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 2);
		}

		#region Funciones Ultragrid
		private void cmdTotal()
		{
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
			int iFiscal = 0;
			if (this.chkFiscal.Checked) iFiscal = 1;
			
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			if ((int)this.txtNumidOrigen.Value > 0)
			{
				dtFecha = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Select Fecha From Compra Where idCompra = {0}", (int)this.txtNumidOrigen.Value));

				dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
					dtFecha.ToString("yyyyMMdd"), iFiscal));
			}
			else if ((int)this.txtNumidOrigen.Value == 0)
			{
				dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
					dtFecha.ToString("yyyyMMdd"), iFiscal));
			}

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round((int)dr.Cells["Cantidad"].Value * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round(((int)dr.Cells["Cantidad"].Value * Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value)/100), 2);
					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{
						this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);						
						this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);						
					}
					else
					{
						this.txtIva0.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);						
						this.txtDescIva0Total.Value = Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
				}
				
				this.txtIvaTotal.Value = (Math.Round(Convert.ToDecimal(this.txtIva.Value), 2) - Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)) * dPorcIva /100.00m;

				this.txtTotal.Value = Convert.ToDecimal(this.txtIva0.Value) 
					+ Convert.ToDecimal(this.txtIva.Value)
					- Convert.ToDecimal(this.txtDescIva0Total.Value)
					- Convert.ToDecimal(this.txtDescIvaTotal.Value)
					+ Convert.ToDecimal(this.txtIvaTotal.Value);
			} 
			#endregion Desplazamiento de grilla
		}

		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}
		#endregion Funciones Ultragrid

		private bool ValidaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			
			int iCantidad = 0;

			if (this.chkConsignacion.Checked)
			{
				string sSQL = string.Format("Exec SumaSaldoConsignacion {0}, {1}", (int)e.Cell.Row.Cells["idArticulo"].Value, (int)this.txtNumIdCliente.Value);
				iCantidad = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			}
			if ((int)this.txtNumidOrigen.Value > 0)
			{
				string sSQL = string.Format("Select Cantidad From DetCompra Where idDetCompra = {0}", (int)e.Cell.Row.Cells["idOrigen"].Value);
				iCantidad = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			}

			if ((int)e.Cell.Value > iCantidad) 
			{
				e.Cell.Value = iCantidad;
				MessageBox.Show(string.Format("La cantidad no puede ser mayor a {0}.", iCantidad), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
		
			return true;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string st = "";
			if (e.Cell.Column.ToString() == "Codigo")
			{
				st = e.Cell.Row.Cells["Codigo"].Value.ToString();
			}
			
			if (!bActualiza) return;
			
			bActualiza = false;
		
			int VerificaCambio = 0;
			#region Codigo o Articulo
			if (e.Cell.Column.ToString() == "Codigo" || e.Cell.Column.ToString() == "Articulo")
			{
				#region Variables
				string stCod = "";
				int IdArt = 0;
				int iCuenta = 0;
				if (bRastreo) Rastreo("Inicio Cambio de " + e.Cell.Column.ToString());
				#endregion Variables
			
				#region Codigo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();					
				}
				#endregion Codigo

				#region Articulo
				if (e.Cell.Column.ToString() == "Articulo")
				{
					stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
				}
				#endregion Articulo

				#region Cuenta coincidencias del codigo
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, 2, 0", stCod, (int)this.txtIdTipoFactura.Value, (int) this.cmbBodega.Value);
				if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				#region Mas de una coincidencia o ninguna muestra la lista
				if (iCuenta != 1)
				{ 
					if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));

					int Bodega = (int) this.cmbBodega.Value;
					using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(stCod, 5))					
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							if (bRastreo) Rastreo("Articulo escogido de la lista");
							IdArt = miBusqueda.idArticulo;
						}
						else
						{
							e.Cell.Row.Cells["idArticulo"].Value = 0;
							e.Cell.Row.Cells["Precio"].Value = 0.0;
							e.Cell.Row.Cells["Codigo"].Value = "";
							e.Cell.Row.Cells["Articulo"].Value = "";
							cmdTotal();

							if (bRastreo) Rastreo("Cancela la selección de Artículos en lista");
							bActualiza = true;
							return;						
						}
					}
				}
					#endregion Mas de una coincidencia o ninguna muestra la lista
					#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
				else 
				{ 
					if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
					stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, (int) this.cmbBodega.Value, (int)this.txtIdTipoFactura.Value);
					IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
				}
				#endregion Encuentra un solo articulo
				#region Asignaciones
				if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
				stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 5, '{2}', {3}", IdArt, idGrupoCliente, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (bool)this.chkFiscal.Checked);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
				dr.Read();
				e.Cell.Row.Cells["Codigo"].Value = dr.GetString(0);
				e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
				e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
				e.Cell.Row.Cells["Impuesto"].Value = dr.GetDecimal(4);
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				dr.Close();
			
				VerificaCambio = 2; // Verificaciones y Cambio de precio
				#endregion Asignaciones
			}
			#endregion Codigo o articulo						

			#region Cantidad
			if (e.Cell.IsInEditMode)
			{ 
				if (e.Cell.Column.ToString() == "Cantidad")
				{					
					if ((int)this.txtNumidOrigen.Value > 0 || this.chkConsignacion.Checked)
					{
						if (!ValidaCantidad(e)) 
						{	
							if (bRastreo) Rastreo("Precio menor al establecido");
							cmdTotal();
							bActualiza = true;
							return;
						}						
					}
					cmdTotal();
				}
				VerificaCambio = 1; 
			}
			#endregion Cantidad			
				
			#region Precio
			if (e.Cell.Column.ToString() == "Precio")
			{				
				cmdTotal();
			}
			#endregion Precio

			#region IVA
			if (e.Cell.Column.ToString() == "Impuesto")
			{				
				bActualiza = false;						
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Seleccione un Producto antes de Ingresar el Porcentaje de I.V.A.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
				{	
					if (!this.chkFiscal.Checked)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Para Notas de Credito No Fiscales El Porcentaje de I.V.A. debe ser 0%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}

					if ((decimal)e.Cell.Row.Cells["Impuesto"].Value != dPorcIva)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show(string.Format("El Porcentaje de I.V.A. debe ser 0% o {0}%", dPorcIva), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				cmdTotal();
				bActualiza = true;
				VerificaCambio = 1; // Verificaciones y cambio de precio conficional		
			}
			#endregion IVA

			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				bActualiza = false;
				cmdTotal();
				bActualiza = true;
			}			
			#endregion DescuentoPorc o DescuentoArt
			bActualiza = true;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 5, (int)this.cmbBodega.Value);
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			int iTotalLineas = ultraGrid1.Rows.Count;
			string stExec = string.Format("Exec CompraNumeroLineas {0}, {1}", (int)this.txtIdTipoFactura.Value, iTotalLineas);
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);

			if (stMensaje.Length > 0) 
			{
				e.Cancel = true;
				MessageBox.Show(stMensaje);
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{
					e.Cancel = true;
				}
				else
				{
					e.DisplayPromptMsg = false;
				}
			}
			else e.DisplayPromptMsg = false;
		}

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true;
				string stError = e.DataErrorInfo.InvalidValue.ToString();
				string stCol = e.DataErrorInfo.Cell.Column.ToString();
				string stMensaje = "Error Grilla: DEVOLUCION EN VENTA";
				int iLong = stMensaje.Length;
				if (iLong > 20) iLong = 20;

				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);

				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol + "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", "Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{
						#region F3
					case (int) Keys.F3:
						#region Inventario					
						#region Busqueda servicios
						if (this.chkConsignacion.Checked)
						{
							#region Articulos
							if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
							{
								using (frmBuscarArticulosLiquidConsignacion miBusqueda = new frmBuscarArticulosLiquidConsignacion((int)this.txtNumIdCliente.Value))
								{
									if (DialogResult.OK == miBusqueda.ShowDialog())
									{		
										ultraGrid1.ActiveRow.Cells["idArticulo"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
										ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
										ultraGrid1.ActiveRow.Cells["Articulo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString();										
										ultraGrid1.ActiveRow.Cells["Cantidad"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Saldo"].Value;										
									}
								}
							}
							#endregion Articulos							
						}
						else
						{											
							if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion)
							{	
								if ((int)this.txtNumidOrigen.Value > 0 && this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) return;
								using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios("", 5))
								{
									if (DialogResult.OK == miBusqueda.ShowDialog())
									{										
										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
									}
								}
							}
						}		
						#endregion Busqueda servicios
						#endregion Inventario
						break;
						#endregion F3										
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al cargar el artículo.");
			}

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}
		#endregion Ultragrid

		private void cmbMotivos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bEdicion = false;
			bNuevo = false;

			this.btnNuevo.Focus();

			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
						
			this.txtNumero.Text = "";
			this.txtComprobante.Text = "";
			this.chkFiscal.Checked = false;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.dtFecha.Value = System.DBNull.Value;
			this.cmbMotivos.Value = System.DBNull.Value;
			this.txtidProyecto.Value = 0;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbBodegaDestino.Value = System.DBNull.Value;

			this.txtNumIdCompra.Value = 0;
			this.txtNumEstado.Value = 0;
			this.txtNumIdAsiento.Value = 0;
			this.txtNumIdCliente.Value = 0;			
			this.txtNumidOrigen.Value = 0;
			this.txtNumIdAsiento.Value = 0;						
			this.txtNumIdproyectoZoom.Value = 0;
			this.txtNumIdFormaPago.Value = 0;
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			this.txtNotas.Text = "";
			this.txtObservaciones.Text = "";
      			
			this.txtNumero.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbMotivos.Enabled = false;			
			this.cmbVendedor.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbBodegaDestino.Enabled = false;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Checked = false;
			this.chkFiscal.Checked = false;
			this.txtNotas.Enabled = false;
			this.txtObservaciones.Enabled = false;		

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAutorizar.Enabled = false;
			this.btnAsiento.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");

			#region Variables
			bImpreso = false;
			bEliminaAlValidar = false;			
			bAnulado = false;
			idBloqueaTransacciones = 0;
			dPorcIva = 0.00m;	
			idGrupoCliente = 0;
			sTipoOperacion = "";
			#endregion Variables

			this.btnNuevo.Focus();			
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			try
			{					
				this.btnCancelar_Click(sender, e);

				#region Tipo de Documento
				using (frmFiscalInterna FI = new frmFiscalInterna(miAcceso.BNCFiscal))
				{				
					if (DialogResult.OK == FI.ShowDialog())
					{
						this.chkFiscal.Checked = (bool)FI.optFiltro.Value;
						this.chkConsignacion.Checked = FI.chkConsignacion.Checked;
						bAnulado = FI.chkAnulado.Checked;
					}
					else 
					{
						this.btnCancelar_Click(sender, e);
						return;
					}
				}
				#endregion Tipo de Documento

				bNuevo = true;

				#region Carga Documento de Origen
				if (!bAnulado)
				{
					using (frmBuscaFacturasCruzaNC Busqueda = new frmBuscaFacturasCruzaNC(1))
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{							
							this.txtNumIdCompra.Value = 0;
							this.txtNumidOrigen.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
							idGrupoCliente = (int)Busqueda.grdFacturas.ActiveRow.Cells["idGrupoCliente"].Value;
							this.txtNumIdFormaPago.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idFormaPago"].Value;
							this.cmbBodega.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["Bodega"].Value;
							this.txtidProyecto.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idProyecto"].Value;
							this.cmbVendedor.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idVendedor"].Value;
							this.txtComprobante.Text = Busqueda.grdFacturas.ActiveRow.Cells["Numero"].Value.ToString();
							this.cmbTipoRuc.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idTipoRuc"].Value;
							this.txtNumIdCliente.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idCliente"].Value;
							this.txtRuc.Text = Busqueda.grdFacturas.ActiveRow.Cells["Ruc"].Value.ToString();
							this.txtNombre.Text = Busqueda.grdFacturas.ActiveRow.Cells["Nombre"].Value.ToString();							
							if (Busqueda.grdFacturas.ActiveRow.Cells["idProyectoZoom"].Value != System.DBNull.Value) this.txtNumIdproyectoZoom.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idProyectoZoom"].Value; 							

							if (this.chkFiscal.Checked)
							{
								this.txtIva0.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["SubtotalExcento"].Value;
								this.txtDescIva0Total.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["Descuento0"].Value;
								this.txtIva.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["SubtotalIva"].Value;
								this.txtDescIvaTotal.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["Descuento"].Value;							
								this.txtIvaTotal.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["Iva"].Value;
								this.txtTotal.Value = (decimal)Busqueda.grdFacturas.ActiveRow.Cells["Total"].Value;
							}
														
							if (this.chkFiscal.Checked)						
								this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePedidosReservasConsignaciones {0}", (int)this.txtNumidOrigen.Value));							
						}
						else 
						{
							this.btnCancelar_Click(sender, e);
							return;
						}
					}
				}
				#endregion Carga Documento de Origen

				#region Controles
				//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;				
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				this.txtNumEstado.Value = 1;
				
				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
				
				if (this.chkFiscal.Checked)
				{
					if (!bNumeroAutomatico) this.txtNumero.Enabled = true;					
					if (this.cmbVendedor.ActiveRow == null && !bAnulado) this.cmbVendedor.Enabled = true;
					if (!bAnulado) this.cmbBodegaDestino.Enabled = true;
					if (bAnulado) this.cmbBodega.Value = 1;
				}
				else
				{	
					this.txtNumIdFormaPago.Value = 6;	
					this.txtNumContadoCredito.Value = 1;
					this.cmbBodega.Value = 1;
					this.cmbTipoRuc.Value = 5;

					this.cmbBodega.Enabled = true;	
					if (this.chkConsignacion.Checked) this.cmbBodegaDestino.Enabled = true;
					this.cmbTipoRuc.Enabled = true;																
					this.txtRuc.Enabled = true;					
				}

				if (!bAnulado) this.cmbMotivos.Enabled = true;

				this.txtNotas.Enabled = true;			

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				if (!bAnulado) this.btnGuardar.Enabled = true;
				this.btnAsiento.Enabled = false;
				this.btnAutorizar.Enabled = false;
				if (bAnulado) this.btnAnular.Enabled = true;
				this.btnCancelar.Enabled = true;
				
				if (!bAnulado) FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);	
				#endregion Controles

				bEdicion = true;
				
				if (miAcceso.CambiarFecha) this.dtFecha.Focus();
				else if (this.chkFiscal.Checked) 
				{
					if (!bAnulado)
						if (this.cmbVendedor.Enabled) this.cmbVendedor.Focus(); else this.cmbBodegaDestino.Focus();
					else
						if (bNumeroAutomatico) this.btnAnular.Focus(); else this.txtNumero.Focus();
				}				
				else if (!this.chkFiscal.Checked) this.cmbMotivos.Focus();			
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Consultar(int idCompra)
		{
			try
			{
				#region Botones
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAutorizar.Enabled = false;
				this.btnAsiento.Enabled = false;
				#endregion Botones

				#region Maestro
				string sSQL = string.Format("Exec ConsultaIndividualNotaDeCreditoCliente {0}", idCompra);
				SqlDataReader drNotaDeCredito = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drNotaDeCredito.Read();
				if (drNotaDeCredito.HasRows)
				{					
					this.txtNumIdCompra.Value = idCompra;
					if (drNotaDeCredito.GetValue(1) != System.DBNull.Value) this.txtNumidOrigen.Value = drNotaDeCredito.GetInt32(1);
					this.txtIdSucursal.Value  = drNotaDeCredito.GetInt32(2);
					this.txtIdTipoFactura.Value = drNotaDeCredito.GetInt32(3);
					this.txtIdComprobante.Value  = drNotaDeCredito.GetInt32(4);
					this.txtNumIdAsiento.Value = drNotaDeCredito.GetInt32(5);
					this.txtNumContadoCredito.Value = drNotaDeCredito.GetInt32(6);
					this.txtNumIdFormaPago.Value = drNotaDeCredito.GetInt32(7);
					this.chkConsignacion.Checked = drNotaDeCredito.GetBoolean(8);
					this.chkFiscal.Checked = drNotaDeCredito.GetBoolean(9);

					this.cmbTipoRuc.Value = drNotaDeCredito.GetInt32(10);
					this.txtNumIdCliente.Value = drNotaDeCredito.GetInt32(11);
					this.txtRuc.Text = drNotaDeCredito.GetValue(12).ToString();
					this.txtNombre.Text = drNotaDeCredito.GetValue(13).ToString();

					this.txtNumEstado.Value = drNotaDeCredito.GetInt32(14);
					this.lblEstado.Text = drNotaDeCredito.GetValue(15).ToString();

					this.cmbBodega.Value = drNotaDeCredito.GetInt32(16);
					this.cmbBodegaDestino.Value = drNotaDeCredito.GetInt32(17);
					this.cmbMotivos.Value = drNotaDeCredito.GetInt32(18);
					this.txtidProyecto.Value = drNotaDeCredito.GetInt32(19);
					if (drNotaDeCredito.GetValue(20) != System.DBNull.Value) this.txtNumIdproyectoZoom.Value = drNotaDeCredito.GetInt32(20);
					this.cmbVendedor.Value = drNotaDeCredito.GetInt32(21);
					this.txtNumero.Text = drNotaDeCredito.GetValue(22).ToString();
					this.txtComprobante.Text = drNotaDeCredito.GetValue(23).ToString();

					this.dtFecha.Value = drNotaDeCredito.GetDateTime(24);
					this.txtSerie.Text = drNotaDeCredito.GetValue(25).ToString();
					this.txtAutFactura.Text = drNotaDeCredito.GetValue(26).ToString();
					if (drNotaDeCredito.GetValue(27) != System.DBNull.Value) this.cmbCaducidad.Value = drNotaDeCredito.GetDateTime(27);

					this.txtIva0.Value = drNotaDeCredito.GetDecimal(28);
					this.txtDescIva0Total.Value = drNotaDeCredito.GetDecimal(29);
					this.txtIva.Value = drNotaDeCredito.GetDecimal(30);
					this.txtDescIvaTotal.Value = drNotaDeCredito.GetDecimal(31);
					this.txtIvaTotal.Value = drNotaDeCredito.GetDecimal(32);
					this.txtTotal.Value = drNotaDeCredito.GetDecimal(33);

					bImpreso = drNotaDeCredito.GetBoolean(34);
					bAnulado = drNotaDeCredito.GetBoolean(35);
					this.lblUsuario.Text = drNotaDeCredito.GetValue(36).ToString();

					this.txtNotas.Text = drNotaDeCredito.GetValue(37).ToString();
					this.txtObservaciones.Text = drNotaDeCredito.GetValue(38).ToString();
					idGrupoCliente = drNotaDeCredito.GetInt32(39);
				}
				drNotaDeCredito.Close();				
				#endregion Maestro

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value));				

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				if (!bAnulado)
				{
					if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.BAutorizar) if ((int)this.txtNumEstado.Value < 8) this.btnAutorizar.Enabled = true;

					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
					{
						if (miAcceso.Editar) 
						{
							if ((int)this.txtNumEstado.Value == 1) this.btnEditar.Enabled = true;
							else if (((int)this.txtNumEstado.Value == 5 || (int)this.txtNumEstado.Value == 7) && miAcceso.BEditarAprobado) this.btnEditar.Enabled = true;            												
						}
						if (miAcceso.Anular && !bAnulado)
						{
							if ((int)this.txtNumEstado.Value == 1) this.btnAnular.Enabled = true;
							else if (((int)this.txtNumEstado.Value == 5 || (int)this.txtNumEstado.Value == 7) && miAcceso.BAnularAprobado) this.btnAnular.Enabled = true;
							else if (((int)this.txtNumEstado.Value == 8) && miAcceso.BAnularLiquidado) this.btnAnular.Enabled = true;
						}
						if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
					}
				}		
				
				this.btnCancelar.Enabled = true;				
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);									
				}
			}		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtNumIdCompra.Value);	

			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid1, miAcceso.Editar, (int)this.txtNumEstado.Value, 5, (int)this.cmbBodegaDestino.Value, 
				(DateTime)this.dtFecha.Value, false, bAnulado, 0, cdsSeteoF, bImpreso, false, miAcceso.BEditarAprobado, false, sTipoOperacion, (int)this.cmbMotivos.Value, this.chkFiscal.Checked)) return;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;			
			
			if ((int)this.cmbMotivos.Value != 6)this.cmbMotivos.Enabled = true;
			
			if (!this.chkFiscal.Checked) this.cmbBodega.Enabled = true;
			if (this.chkFiscal.Checked) this.cmbBodegaDestino.Enabled = true;
			this.txtNotas.Enabled = true;
			
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAutorizar.Enabled = false;
			this.btnAsiento.Enabled = false;			
			this.btnCancelar.Enabled = true;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);				

			bEdicion = true;
			bNuevo = false;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnMotivo_Click(object sender, System.EventArgs e)
		{
			using (frmMotivosDevolucion miMotivo = new frmMotivosDevolucion(5))
			{
				if (DialogResult.OK == miMotivo.ShowDialog())
				{					
					this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 Order By Motivo");
					this.cmbMotivos.DisplayLayout.Bands[0].Columns["Motivo"].Width = 300;
					this.cmbMotivos.DisplayLayout.Bands[0].Columns["idMotivo"].Hidden = true;	
				}
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{			
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, sTipoOperacion))
			{
				string sMensaje = "ANULAR";
				if (bAnulado) sMensaje = "ELIMINAR";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede '{1}' esta Nota de Credito, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper(), sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}
			
//			if (bNuevo)
//			{				
//				if (this.chkFiscal.Checked)
//					if (!Validacion.vbNumeracion(bNumeroAutomatico, (int)this.txtNumIdCompra.Value, 5, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
//
//				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta(20))
//				{
//					if (DialogResult.OK == miNota.ShowDialog())
//					{						
//						#region Graba Nota De Credito
//						#region Maestro
//						int IBodegaDestino = 0;
//						if (this.cmbBodegaDestino.Value != System.DBNull.Value) IBodegaDestino = (int)this.cmbBodegaDestino.Value;
//
//						#region Fechas
//						DateTime dtFecha = (DateTime)this.dtFecha.Value;
//						DateTime dtFechaCad = DateTime.Today;
//						DateTime dtFechaEnt = DateTime.Today;
//						DateTime dtHoraEntrega = DateTime.Today;
//						DateTime dtFechPrimPago = DateTime.Today;
//						bool bFecCad = false;
//						bool bFecEnt = false;
//						bool bHorEnt = false;
//						bool bFecPrimPag = false;
//						if (this.cmbCaducidad.Value != System.DBNull.Value) 
//						{
//							bFecCad = true;
//							dtFechaCad = (DateTime) this.cmbCaducidad.Value;
//						}
//						#endregion Fechas
//
//						this.txtNumIdCompra.Value = 
//							Facturacion.GrabaFacturacion((int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtIdSucursal.Value, 
//							(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
//							(int)this.txtNumEstado.Value,
//							(bool)this.chkConsignacion.Checked,
//							(int)this.txtNumContadoCredito.Value, (int)this.txtNumIdFormaPago.Value,
//							(int)this.cmbBodega.Value, (int)this.txtidProyecto.Value, 0,
//							this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
//							0, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
//							(DateTime)this.dtFecha.Value, 
//							this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad, 
//							false, DateTime.Today, DateTime.Today, 0,
//							this.txtNotas.Text.ToString(),
//							false, 0.00m, 0, 0.00m, DateTime.Today,
//							(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
//							false, bFecCad, false, false, true,
//							(bool)this.chkFiscal.Checked, IBodegaDestino, 0, this.txtObservaciones.Text.ToString(), (decimal)this.txtTotal.Value, 0.00m, 0.00m, 0.00m, 0.00m, false, 0.00m, 0, 0, 0, 0, 0, "", false,
//							1, "", "", 0, "", 0.00m, 0, "", 0,
//							bNuevo, bNumeroAutomatico, idBloqueaTransacciones, this.ultraGrid1, false, 1, this.ultraGrid1, false);		
//						#endregion Maestro
//						#endregion Graba Nota De Credito
//										
//						#region Controles
//						bEliminaAlValidar = false;
//										
//						this.cmbTipoRuc.Enabled = false;
//						this.txtRuc.Enabled = false;
//						this.txtNumero.Enabled = false;
//						this.dtFecha.Enabled = false;
//						this.cmbMotivos.Enabled = false;							
//						this.cmbVendedor.Enabled = false;
//						this.cmbBodega.Enabled = false;
//						this.cmbBodegaDestino.Enabled = false;
//						this.txtNotas.Enabled = false;
//						this.txtObservaciones.Enabled = false;
//						this.chkFiscal.Enabled = false;
//						bAnulado = true;
//						this.lblEstado.Text = "ANULADO";
//						
//						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
//
//						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
//						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;							
//						if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
//						else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
//						this.btnGuardar.Enabled = false;
//						if (miAcceso.BAutorizar) this.btnAutorizar.Enabled = true;							
//						this.btnAnular.Enabled = false;
//
//						this.btnCancelar.Enabled = true;					
//
//						bNuevo = false;
//						bEdicion = false;	
//						#endregion Controles
//					}
//				}
//				
//				return;
//			}
			
			if (!bNuevo)
			{
				this.Consultar((int)this.txtNumIdCompra.Value);

				if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid1, miAcceso.Anular, (int)this.txtNumEstado.Value, 5, (int)this.cmbBodegaDestino.Value, 
					(DateTime)this.dtFecha.Value, true, bAnulado, 0, cdsSeteoF, bImpreso, false, miAcceso.BAnularAprobado, miAcceso.BAnularLiquidado, sTipoOperacion, (int)this.cmbMotivos.Value, this.chkFiscal.Checked)) return;

				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));

				if ((iEstado == 5 || iEstado == 7) && !miAcceso.BAnularAprobado)
				{
					MessageBox.Show("No puede ANULAR Notas de Crédito Aprobadas \n\n Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				}
				else if (iEstado == 8 && !miAcceso.BAnularLiquidado)
				{
					MessageBox.Show("No puede anular NOTA DE CRÉDITO CANCELADA \n\n Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				}
				else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))			
				{
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular Esta Nota De Credito", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{
								try
								{
									string sSQLAnul = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", (int)this.txtNumIdCompra.Value, (int)this.txtIdTipoFactura.Value, Nota.txtNotas.Text.ToString());
									Funcion.sEscalarSQL(cdsSeteoF, sSQLAnul);

									bNuevo = false;
									bEdicion = false;
									
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

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtNumIdCompra.Value);

			Facturacion.Impresion((int)this.txtNumIdCompra.Value, bAnulado, (int)this.txtNumEstado.Value, miAcceso.Imprimir, miAcceso.ReImprimir, this.chkFiscal.Checked, bImpreso, 5, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validaciones
			if (this.chkFiscal.Checked)
			{
				if ((int)this.txtNumidOrigen.Value == 0)
				{
					MessageBox.Show("Nota de Credito Fiscal debe venir de una Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}
			}

			if (this.chkFiscal.Checked && !Validacion.vbValidaValorDctoOrigen((int)this.txtNumidOrigen.Value, Convert.ToDecimal(this.txtTotal.Value), cdsSeteoF)) return;
					
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha", true, sTipoOperacion, cdsSeteoF)) return; 

			if (!Validacion.vbClienteExiste(this.txtRuc, false, 5, (int)this.txtNumIdFormaPago.Value, cdsSeteoF)) return;		
			
			if (this.chkFiscal.Checked)
				if (!Validacion.vbNumeracion(bNumeroAutomatico, (int)this.txtNumIdCompra.Value, 5, (int)this.cmbBodega.Value, this.txtNumero, cdsSeteoF)) return;
									
			if (!Validacion.vbComboVacio(this.cmbMotivos, "Seleccione el motivo")) return;
						
			if ((int)this.txtNumidOrigen.Value > 0 || this.chkConsignacion.Checked)
				if (this.chkFiscal.Checked)
					if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccione la Bodega de Destino")) return;
						
			if (this.chkFiscal.Checked)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el vendedor")) return;
			
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 5)) return;

			cmdTotal();

			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			if (!Validacion.vbTexto(this.txtNotas, 20, 200, "Notas")) return;
			#endregion Validaciones
			
			try
			{	
				#region Graba Nota De Credito
				int IBodegaDestino = 0;
				int idVendedor = 0;
				bool bFecCad = false;
				DateTime dtFechaCad = DateTime.Today;
				if (this.cmbBodegaDestino.ActiveRow != null) IBodegaDestino = (int)this.cmbBodegaDestino.Value;
				if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;								
				if (this.cmbCaducidad.Value != System.DBNull.Value) 
				{
					bFecCad = true;
					dtFechaCad = (DateTime) this.cmbCaducidad.Value;
				}										

				if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACIÓN DEL CLIENTE \n PRODUCTOS \n PRECIOS CANTIDADES DESCUENTOS \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					this.txtNumIdCompra.Value = 
						Facturacion.GrabaFacturacion((int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtIdSucursal.Value, 
						(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
						(int)this.txtNumEstado.Value,
						(bool)this.chkConsignacion.Checked,
						(int)this.txtNumContadoCredito.Value, (int)this.txtNumIdFormaPago.Value,
						(int)this.cmbBodega.Value, (int)this.txtidProyecto.Value, 
						idVendedor,
						this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
						(int)this.cmbTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
						(DateTime)this.dtFecha.Value, 
						this.txtSerie.Text.ToString(), this.txtAutFactura.Text.ToString(), dtFechaCad, 
						false, DateTime.Today, DateTime.Today, 0,
						this.txtNotas.Text.ToString(),
						false, 0.00m, 0, 0.00m, DateTime.Today,
						(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value,
						false, bFecCad, false, false, bAnulado,
						(bool)this.chkFiscal.Checked, IBodegaDestino, (int)this.cmbMotivos.Value, this.txtObservaciones.Text.ToString(), (decimal)this.txtTotal.Value, 0.00m, 0.00m, 0.00m, 0.00m, false, 0.00m, 0, 0, 0, 0, 0, "", false,
						1, "", "", 0, "", 0.00m, 0, "", 0.00m,
						bNuevo, bNumeroAutomatico, idBloqueaTransacciones, this.ultraGrid1, false, 1, this.ultraGrid1, false);
	
//					if (bNuevo)
//					{
						string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 1", (int)this.txtNumIdCompra.Value));

						string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
						string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
						string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, (int)this.txtNumIdCompra.Value));
//					}
				}			
				#endregion Graba Nota De Credito		

				#region Controles
				if ((int)this.txtNumIdCompra.Value == 0) return;

				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false;
				this.txtNumero.Enabled = false;
				this.dtFecha.Enabled = false;
				this.cmbMotivos.Enabled = false;					
				this.cmbVendedor.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbBodegaDestino.Enabled = false;
				this.txtNotas.Enabled = false;
				this.txtObservaciones.Enabled = false;
				this.chkFiscal.Enabled = false;

				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
					
				bNuevo = false;
				bEdicion = false;
				idBloqueaTransacciones = 0;

				this.Consultar((int)this.txtNumIdCompra.Value);
				#endregion Controles
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
			}
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 5, this.txtRuc.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtNumIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[2] != null) idGrupoCliente = int.Parse(sParametros[2]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];			
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtNumIdCliente.Value = 0;				
				this.txtNombre.Text = "";
				idGrupoCliente = 0;

				CargaCliente();
			}
		}

		private void cmbMotivos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				if (!this.chkFiscal.Checked) this.txtRuc.Focus();
				else if (!bNumeroAutomatico) this.txtNumero.Focus();
				else
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
			}
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
				this.cmbMotivos.Focus();
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
				if (this.chkFiscal.Checked) this.cmbBodegaDestino.Focus(); else this.cmbMotivos.Focus();
		}

		private void cmbBodegaDestino_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
				if (this.chkFiscal.Checked) 
					if (this.cmbVendedor.Enabled) this.cmbVendedor.Focus(); else this.cmbMotivos.Focus();
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObservaciones.Focus();
		}

		private void txtObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnGuardar.Focus();
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			if (bEdicion)
			{
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
								
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 5, (int)this.txtNumIdFormaPago.Value, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
			}
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) e.Cancel = true;
		}

		private void cmbMotivos_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
				if (!Validacion.vbComboVacio(this.cmbMotivos, "Seleccione un motivo")) e.Cancel = true;
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtNumIdCompra.Value);

			FuncionesContabilidad.AsientoDocumentosCompra((int)this.txtNumIdAsiento.Value, (int)this.txtNumIdCompra.Value, 5, miAcceso.CrearAsiento, bAnulado, (DateTime)this.dtFecha.Value, cdsSeteoF);			
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			if (!miAcceso.BAutorizar)
			{
				MessageBox.Show("No tiene Acceso a Autorizar Notas de Crédito", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else
			{
				this.Consultar((int)this.txtNumIdCompra.Value);

				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));

				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Borrar From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true))
				{
					MessageBox.Show("La Nota de Credito esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				}				
				else if (iEstado == 8)
				{
					MessageBox.Show("La Nota de Credito esta cancelada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				}					
				else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))			
				{
					using (frmAutNotaDeCredito miAutorizacion = new frmAutNotaDeCredito((int)this.txtNumEstado.Value, this.txtObservaciones.Text.ToString()))
					{
						if (DialogResult.OK == miAutorizacion.ShowDialog())
						{
							this.txtNumEstado.Value = miAutorizacion.Estado;
							string sSQL = string.Format("Update Compra Set Estado = {0}, Observaciones = '{1}' Where idCompra = {2}", miAutorizacion.Estado, miAutorizacion.txtNotas.Text.ToString(), (int)this.txtNumIdCompra.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							if ((int)miAutorizacion.Estado == 4) this.lblEstado.Text = "RECHAZADO";	else if ((int)miAutorizacion.Estado == 5) this.lblEstado.Text = "AUTORIZADO";					
						}
					}
				}
			}
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				this.Consultar((int)this.txtNumIdCompra.Value);

				FuncionesContabilidad.ValidaVistaAsientoDesdeDocumentos((int)this.txtNumIdAsiento.Value, bAnulado, miAcceso.CrearAsiento, cdsSeteoF);
			}			
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
			if (e.KeyChar == 13)
			{
				if (this.cmbBodega.Enabled)	this.cmbBodega.Focus();
				else if (this.chkFiscal.Checked) this.cmbBodegaDestino.Focus(); else this.cmbMotivos.Focus();	
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			if (bEdicion)
				cmdTotal();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalVendedor {0}", (int)this.cmbBodega.Value));
				if (bEdicion) this.txtidProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
			
				#region S.R.I.
				if (bNuevo && (bool)this.chkFiscal.Checked)
				{
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select AutNotaDeCredito, SerieNotaDeCredito, CaducaNotaDeCredito From Seteo", true);	
					dr.Read();
					if (dr.HasRows)
					{
						this.txtAutFactura.Text = dr.GetValue(0).ToString();
						this.txtSerie.Text = dr.GetValue(1).ToString();
						if (dr.GetValue(2) != System.DBNull.Value) this.cmbCaducidad.Value = dr.GetDateTime(2);
					}
					dr.Close();
				}
				#endregion S.R.I.
			}
		}

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegaDestino_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Local")) e.Cancel = true;
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value && (int)this.txtNumidOrigen.Value > 0) e.Cancel = true;
			}

//			if (e.Cell.Column.ToString() == "DescuentoPorc")
//			{
//				if (this.chkConsignacion.Checked) e.Cancel = true;				
//			}
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.Length > 0)
			{
				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);
			}
		}

		private void btnFe_Click(object sender, System.EventArgs e)
		{
			return;
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
			string sDocumentoModificado = this.txtComprobante.Text.ToString();
			DateTime dtFechaDocOrigen = DateTime.Today;
			decimal dTotalDocModificado = 0;
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
			string stArchivo = "NCPRUEBA.XML";
			XmlTextWriter XMLFactura = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));
			XMLFactura.Formatting = Formatting.Indented;
			XMLFactura.WriteStartDocument(true);
			XMLFactura.WriteStartElement("notaCredito");
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

			#region Informacion Nota De Crédito
			sDocumentoModificado = sEstab + "-" + sPtoEmi + "-" + sDocumentoModificado;

			string sSQLInfoDocMod = string.Format("Select Fecha, Total From Compra Where idCompra = {0} And idTipoFactura = 1", (int)this.txtNumidOrigen.Value);
			SqlDataReader drInfoDocMod = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfoDocMod, true);
			drInfoDocMod.Read();
			if (drInfoDocMod.HasRows)
			{					 
				dtFechaDocOrigen = (DateTime)drInfoDocMod.GetValue(0);
				dTotalDocModificado = (decimal)drInfoDocMod.GetValue(1);				
			}
			drInfoDocMod.Close();

			XMLFactura.WriteStartElement("infoNotaCredito");

			XMLFactura.WriteElementString("fechaEmision", dtFecha.ToString("dd/MM/yyyy"));
			XMLFactura.WriteElementString("dirEstablecimiento", sDirEstablecimiento);
			XMLFactura.WriteElementString("tipoIdentificacionComprador", sCodDocClte);
			XMLFactura.WriteElementString("razonSocialComprador", this.txtNombre.Text.ToString());
			XMLFactura.WriteElementString("identificacionComprador", this.txtRuc.Text.ToString());
			XMLFactura.WriteElementString("contribuyenteEspecial", iResolucion.ToString());
			XMLFactura.WriteElementString("obligadoContabilidad", "SI");
			XMLFactura.WriteElementString("codDocModificado", "01");
			XMLFactura.WriteElementString("numDocModificado", sDocumentoModificado);
			XMLFactura.WriteElementString("fechaEmisionDocSustento", dtFechaDocOrigen.ToString("dd/MM/yyyy"));			
			XMLFactura.WriteElementString("totalSinImpuestos", dTotalDocModificado.ToString("0.00", us));
			XMLFactura.WriteElementString("valorModificacion", dTotal.ToString("0.00", us));
			XMLFactura.WriteElementString("moneda", "DOLAR");

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
      
			XMLFactura.WriteElementString("motivo", this.cmbMotivos.Text.ToString());
			
			XMLFactura.WriteEndElement();
			#endregion Informacion Nota De Crédito

			#region Detalles
			XMLFactura.WriteStartElement("Detalles");
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				XMLFactura.WriteStartElement("Detalle");
	
				decimal dCantidad = (decimal)dr.Cells["Cantidad"].Value;
				decimal dPrecio = (decimal)dr.Cells["Precio"].Value;
				decimal dDescuentoUnit = (decimal)dr.Cells["DescuentoArt"].Value;
				decimal dPrecioTotal = (decimal)dr.Cells["Subtotal"].Value - (decimal)dr.Cells["DescuentoArt"].Value;

				XMLFactura.WriteElementString("codigoInterno", dr.Cells["Codigo"].Value.ToString());
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
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtSerie_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAutFactura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCaducidad_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBodegaDestino_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccione una Bodega Destino")) e.Cancel = true;
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, sTipoOperacion, cdsSeteoF)) e.Cancel = true;
		}

		private void frmDevolucionVenta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Nota De Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmDevolucionVenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if ((int)this.txtNumEstado.Value == 1) this.btnEditar_Click(sender, e); 
																											else if (((int)this.txtNumEstado.Value == 5 || (int)this.txtNumEstado.Value == 7) && miAcceso.BEditarAprobado) this.btnEditar_Click(sender, e);																													
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir && !bImpreso) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if ((int)this.txtNumEstado.Value == 1) this.btnAnular_Click(sender, e); 
																																	 else if (((int)this.txtNumEstado.Value == 5 || (int)this.txtNumEstado.Value == 7) && miAcceso.BAnularAprobado) this.btnAnular_Click(sender, e);
																																	 else if (((int)this.txtNumEstado.Value == 8) && miAcceso.BAnularLiquidado) this.btnAnular_Click(sender, e);
				if (e.KeyCode == Keys.F5)
				{
					miAcceso = new Acceso(cdsSeteoF, "0725");
					this.Consultar((int)this.txtNumIdCompra.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);								
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
				dPorcIva = Facturacion.dPorcentajeIVAVigente(this.chkFiscal.Checked, true, (int)this.txtNumidOrigen.Value, (DateTime)this.dtFecha.Value, cdsSeteoF);				
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtRuc.Enabled = true;
				this.txtNumIdCliente.Value = 0;
				this.txtRuc.Value = "";
				this.txtNombre.Value = "";

				if (this.cmbTipoRuc.ActiveRow != null)
				{
					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;					
				}
			}
		}

		private void cmbTipoRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbTipoRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtRuc.Focus();
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkFiscal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkFiscal.Checked) sTipoOperacion = "Venta";
				else sTipoOperacion = "Cobros";
			}
		}
	}
}

