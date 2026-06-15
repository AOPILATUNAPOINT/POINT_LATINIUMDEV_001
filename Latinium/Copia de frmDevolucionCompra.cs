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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDevolucionCompra.
	/// </summary>
	public class frmDevolucionCompra : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.GroupBox groupBox1;

		#region Variables	
		bool bNuevo = false;
		bool bEdicion = false;
		bool bIncluyeIva = false;
		private bool bActualiza = true;
		bool bRastreo = false;
		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;
		bool bAsientoAutomatico = false;
		int IdBusqueda = 0;
		private int IdSigno = 2;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		bool bAutoImpresor = false;
		bool bAnulado = false;
		bool bImpreso = false;
		bool bConsignacion = false;
		#endregion Variables

		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox gbBotones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdproyectoZoom;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCredTributario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFiscal;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;

		[DllImport("user32.dll")]
		static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

		private System.ComponentModel.IContainer components;

		public frmDevolucionCompra()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idCompra = 0;

		public frmDevolucionCompra(int IdCompra)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idCompra = IdCompra;
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDevolucionCompra));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.txtNumIdFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdproyectoZoom = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCredTributario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkFiscal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
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
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(800, 104);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 197;
			// 
			// txtComprobante
			// 
			this.txtComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance1;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(888, 80);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(120, 22);
			this.txtComprobante.TabIndex = 6;
			this.txtComprobante.ValueChanged += new System.EventHandler(this.txtComprobante_ValueChanged);
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(800, 83);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 212;
			this.lblComprobante.Text = "Comprobante";
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 288);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 211;
			this.label2.Text = "Notas";
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 312);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(400, 96);
			this.txtNotas.TabIndex = 8;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 59);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(86, 16);
			this.label8.TabIndex = 233;
			this.label8.Text = "N. Identificación";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 107);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 231;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(103, 16);
			this.label1.TabIndex = 228;
			this.label1.Text = "Buscar Factura (F3)";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 83);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(29, 16);
			this.label5.TabIndex = 226;
			this.label5.Text = "Serie";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(256, 59);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(55, 16);
			this.lblCliente.TabIndex = 224;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(800, 59);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 223;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(112, 104);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 600;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// txtBuscar
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance4;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Enabled = false;
			this.txtBuscar.Location = new System.Drawing.Point(112, 9);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(112, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// txtRuc
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance5;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(112, 56);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 221;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance6;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(328, 56);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(408, 22);
			this.txtNombre.TabIndex = 222;
			// 
			// txtNumero
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance7;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(328, 80);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 500;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtSerieFact
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance8;
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Enabled = false;
			this.txtSerieFact.Location = new System.Drawing.Point(112, 80);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(112, 22);
			this.txtSerieFact.TabIndex = 100;
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			this.txtSerieFact.ValueChanged += new System.EventHandler(this.txtSerieFact_ValueChanged);
			// 
			// txtAutFactura
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance9;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(528, 80);
			this.txtAutFactura.MaxLength = 49;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(208, 22);
			this.txtAutFactura.TabIndex = 598;
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			this.txtAutFactura.Validating += new System.ComponentModel.CancelEventHandler(this.txtAutFactura_Validating);
			this.txtAutFactura.ValueChanged += new System.EventHandler(this.txtAutFactura_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(448, 83);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 235;
			this.label31.Text = "Autorización";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1020, 5);
			this.groupBox1.TabIndex = 237;
			this.groupBox1.TabStop = false;
			// 
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto1.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(720, 160);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(208, 112);
			this.cmbProyecto1.TabIndex = 247;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbBodega
			// 
			this.cmbBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(888, 56);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(120, 21);
			this.cmbBodega.TabIndex = 601;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(760, 315);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(45, 16);
			this.lblSubTotal.TabIndex = 607;
			this.lblSubTotal.Text = "Subtotal";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(760, 363);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 606;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(760, 339);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(57, 16);
			this.lblDescuento.TabIndex = 605;
			this.lblDescuento.Text = "Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(984, 288);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 604;
			this.lblIvaSubt.Text = "IVA";
			this.lblIvaSubt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(880, 288);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 603;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance10;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(840, 384);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(168, 34);
			this.txtTotal.TabIndex = 613;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance11;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(928, 360);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 612;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance12;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(840, 336);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 611;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance13;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(928, 336);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 610;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance14;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(928, 312);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 609;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance15;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(840, 312);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 608;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(760, 393);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 614;
			this.label3.Text = "Total";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance16.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance16;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn7.Header.VisiblePosition = 12;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.Caption = "Artículo";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 194;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 310;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance17;
			ultraGridColumn12.Format = "";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.MaskInput = "";
			ultraGridColumn12.NullText = "0";
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 70;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance18;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 80;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance19;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "IVA";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 40;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance20;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Descuento";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 14;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance21;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "% Desc.";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 65;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance22;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.Header.Caption = "Centro de Costo";
			ultraGridColumn18.Header.VisiblePosition = 16;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn18.Width = 140;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Hidden = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance23;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.MaskInput = "nnnnnnnnn";
			ultraGridColumn21.Width = 53;
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.Caption = "Unidad";
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn23.Width = 53;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn23});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance25;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8.5F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance26;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1000, 144);
			this.ultraGrid1.TabIndex = 615;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
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
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn20});
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 448);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 623;
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
			this.btnConsultar.Location = new System.Drawing.Point(84, 448);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 622;
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
			this.btnEditar.Location = new System.Drawing.Point(166, 448);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 621;
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
			this.btnImprimir.Location = new System.Drawing.Point(246, 448);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 620;
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
			this.btnGuardar.Location = new System.Drawing.Point(327, 448);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 619;
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
			this.btnAnular.Location = new System.Drawing.Point(491, 448);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 618;
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
			this.btnCancelar.Location = new System.Drawing.Point(571, 448);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 617;
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
			this.btnSalir.Location = new System.Drawing.Point(647, 448);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 616;
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
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 432);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1032, 8);
			this.gbBotones.TabIndex = 624;
			this.gbBotones.TabStop = false;
			// 
			// txtNumIdFormaPago
			// 
			this.txtNumIdFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdFormaPago.Enabled = false;
			this.txtNumIdFormaPago.Location = new System.Drawing.Point(968, 448);
			this.txtNumIdFormaPago.Name = "txtNumIdFormaPago";
			this.txtNumIdFormaPago.PromptChar = ' ';
			this.txtNumIdFormaPago.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdFormaPago.TabIndex = 637;
			this.txtNumIdFormaPago.Visible = false;
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(936, 448);
			this.txtNumIdProyecto.Name = "txtNumIdProyecto";
			this.txtNumIdProyecto.PromptChar = ' ';
			this.txtNumIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdProyecto.TabIndex = 635;
			this.txtNumIdProyecto.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(920, 448);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 634;
			this.txtIdComprobante.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(904, 448);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 633;
			this.txtNumEstado.Visible = false;
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(888, 448);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 632;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(872, 448);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 630;
			this.txtIdSucursal.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(856, 448);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 629;
			this.txtNumContadoCredito.Visible = false;
			// 
			// txtNumidOrigen
			// 
			this.txtNumidOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen.Enabled = false;
			this.txtNumidOrigen.Location = new System.Drawing.Point(808, 448);
			this.txtNumidOrigen.Name = "txtNumidOrigen";
			this.txtNumidOrigen.PromptChar = ' ';
			this.txtNumidOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen.TabIndex = 628;
			this.txtNumidOrigen.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(792, 448);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 627;
			this.txtNumIdCompra.Visible = false;
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(824, 448);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 626;
			this.txtNumIdCliente.Visible = false;
			// 
			// txtNumIdAsiento
			// 
			this.txtNumIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdAsiento.Enabled = false;
			this.txtNumIdAsiento.Location = new System.Drawing.Point(840, 448);
			this.txtNumIdAsiento.Name = "txtNumIdAsiento";
			this.txtNumIdAsiento.PromptChar = ' ';
			this.txtNumIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdAsiento.TabIndex = 625;
			this.txtNumIdAsiento.Visible = false;
			// 
			// txtNumIdproyectoZoom
			// 
			this.txtNumIdproyectoZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance30.ForeColor = System.Drawing.Color.Black;
			this.txtNumIdproyectoZoom.Appearance = appearance30;
			this.txtNumIdproyectoZoom.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtNumIdproyectoZoom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdproyectoZoom.Enabled = false;
			this.txtNumIdproyectoZoom.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtNumIdproyectoZoom.ForeColor = System.Drawing.Color.Black;
			this.txtNumIdproyectoZoom.Location = new System.Drawing.Point(952, 448);
			this.txtNumIdproyectoZoom.Name = "txtNumIdproyectoZoom";
			this.txtNumIdproyectoZoom.PromptChar = ' ';
			this.txtNumIdproyectoZoom.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdproyectoZoom.TabIndex = 636;
			this.txtNumIdproyectoZoom.Visible = false;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(407, 448);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 638;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(984, 448);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoRuc.TabIndex = 639;
			this.txtIdTipoRuc.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(256, 83);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 640;
			this.label4.Text = "Número";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(256, 107);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 641;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(1000, 448);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdGrupoCliente.TabIndex = 642;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// txtIdCredTributario
			// 
			this.txtIdCredTributario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCredTributario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCredTributario.Enabled = false;
			this.txtIdCredTributario.Location = new System.Drawing.Point(776, 448);
			this.txtIdCredTributario.Name = "txtIdCredTributario";
			this.txtIdCredTributario.PromptChar = ' ';
			this.txtIdCredTributario.Size = new System.Drawing.Size(16, 22);
			this.txtIdCredTributario.TabIndex = 643;
			this.txtIdCredTributario.Visible = false;
			// 
			// txtIdTipoFormaPago
			// 
			this.txtIdTipoFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFormaPago.Enabled = false;
			this.txtIdTipoFormaPago.Location = new System.Drawing.Point(760, 448);
			this.txtIdTipoFormaPago.Name = "txtIdTipoFormaPago";
			this.txtIdTipoFormaPago.PromptChar = ' ';
			this.txtIdTipoFormaPago.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFormaPago.TabIndex = 644;
			this.txtIdTipoFormaPago.Visible = false;
			// 
			// chkFiscal
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFiscal.Appearance = appearance31;
			this.chkFiscal.BackColor = System.Drawing.Color.Transparent;
			this.chkFiscal.CausesValidation = false;
			this.chkFiscal.Enabled = false;
			this.chkFiscal.Location = new System.Drawing.Point(248, 8);
			this.chkFiscal.Name = "chkFiscal";
			this.chkFiscal.Size = new System.Drawing.Size(48, 22);
			this.chkFiscal.TabIndex = 645;
			this.chkFiscal.Text = "Fiscal";
			this.chkFiscal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkFiscal_KeyDown);
			this.chkFiscal.CheckedChanged += new System.EventHandler(this.chkFiscal_CheckedChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// frmDevolucionCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1016, 478);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.chkFiscal);
			this.Controls.Add(this.txtIdTipoFormaPago);
			this.Controls.Add(this.txtIdCredTributario);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.txtNumIdFormaPago);
			this.Controls.Add(this.txtNumIdproyectoZoom);
			this.Controls.Add(this.txtNumIdProyecto);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumidOrigen);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtNumIdAsiento);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.lblEstado);
			this.KeyPreview = true;
			this.Name = "frmDevolucionCompra";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nota de Crédito Proveedor";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDevolucionCompra_KeyDown);
			this.Resize += new System.EventHandler(this.frmDevolucionCompra_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDevolucionCompra_Closing);
			this.Load += new System.EventHandler(this.frmDevolucionCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Funciones
		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
	
		#region Seguridad Configuracion
		private void UnloadMe()
		{
			this.Close();
		}
		private void Seguridades()
		{
			#region Codigo de Seguridad
			string stTipo = "0612";
			string stPermiso = "61";
								
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			#endregion Tipo de Factura
			if (!Funcion.Permiso(stPermiso, cdsSeteo))
			{
				MessageBox.Show("No puede ingresar a Nota de Crédito Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		} 

		private void Configuracion()
		{
			#region Asigna Signo
			string stSelect = "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 6";
			IdSigno = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
			#endregion Asigna Signo			
			#region Lectura de Archivo - Propiedades - Configuracion
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#endregion Lectura de Archivo - Propiedades - Configuracion
		}
		#endregion Seguridad Configuracion
		#region grid
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
				dDescuentoLineal = 0;
				dIvaLineal = 0;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal( dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value)) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value)/100, 2);
					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0)
					{
						this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal( dr.Cells["Subtotal"].Value), 2);
						this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal( dr.Cells["DescuentoArt"].Value), 2);
					}
					else
					{
						this.txtIva0.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal( dr.Cells["Subtotal"].Value), 2);
						this.txtDescIva0Total.Value = Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
				}
				
				this.txtIvaTotal.Value = Math.Round((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva /100, 2);

				this.txtTotal.Value = Convert.ToDecimal(this.txtIva0.Value)
					+ Convert.ToDecimal(this.txtIva.Value)
					- Convert.ToDecimal(this.txtDescIva0Total.Value)
					- Convert.ToDecimal(this.txtDescIvaTotal.Value)
					+ Convert.ToDecimal(this.txtIvaTotal.Value);
			} 
			#endregion Desplazamiento de grilla
		}

		private void CambiaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((bool) drSeteoF["VerUnidad"] == false) return;
			int IdUnidad = (int) e.Cell.Row.Cells["idUnidad"].Value;
			int IdArticulo = (int) e.Cell.Row.Cells["idArticulo"].Value;
			decimal dUnidades = 0;
			if (e.Cell.Row.Cells["Unidades"].Value != DBNull.Value)
				dUnidades = (decimal) e.Cell.Row.Cells["Unidades"].Value;
			else
				e.Cell.Row.Cells["Unidades"].Value = 0;
		
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVUNP'")) 				
			{
				string stExec = string.Format("Exec ArticuloUnidadCambio {0}, {1}", IdArticulo, IdUnidad);
				e.Cell.Row.Cells["Cantidad"].Value = Funcion.decEscalarSQL(cdsSeteoF, stExec) * dUnidades;
				return;
			}
			string stSelect = "Exec UnidadCambio "	+ IdArticulo;
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsSeteoF, stSelect, false);
			miReader.Read();

			int IdUnidadOr = miReader.GetInt32(0);
			//Poner solo cuando hay unidades visible
			if (IdUnidad == IdUnidadOr)
			{
				decimal dCant = miReader.GetDecimal(4);
				e.Cell.Row.Cells["Cantidad"].Value = (decimal) e.Cell.Row.Cells["Unidades"].Value * dCant;
			}
			else
			{
				decimal dCantidad = miReader.GetDecimal(1); // Unidad Alterna
				if (dCantidad > 0)
				{
					e.Cell.Row.Cells["Cantidad"].Value = 
						(decimal) e.Cell.Row.Cells["Unidades"].Value * dCantidad;
					bool bPeso = miReader.GetBoolean(5);
					if (bPeso) // Venta por peso cambia el precio
					{
						e.Cell.Row.Cells["Precio"].Value = (decimal) e.Cell.Row.Cells["Precio"].Value / dCantidad;
						e.Cell.Row.Cells["Precio"].Refresh();
					}
				}
				if (dCantidad == 0)
				{
					e.Cell.Row.Cells["Cantidad"].Value = (decimal) e.Cell.Row.Cells["Unidades"].Value;
				}
			}
			miReader.Close();
		}

		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}
		private bool MostrarIce = false;
		private bool Limites(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Verifica valores nulos de cantidad o de idArt
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			#endregion Verifica valores nulos de cantidad o de idArt

			#region Lee Iva
			string stIdArt = e.Cell.Row.Cells["idArticulo"].Value.ToString();
			string stExec = string.Format("Exec ArticuloIva {0}, {1}, {2}", 	(int)this.txtIdTipoFactura.Value, stIdArt, 0);
			decimal fIva = Funcion.decEscalarSQL(cdsSeteoF, stExec);
			e.Cell.Row.Cells["Impuesto"].Value = fIva;
			
			if (MostrarIce)
			{
				stExec = string.Format("Exec ArticuloIce {0}, {1}", (int)this.txtIdTipoFactura.Value, stIdArt);
				decimal fIce = Funcion.decEscalarSQL(cdsSeteoF, stExec);
				e.Cell.Row.Cells["IceL"].Value = fIce;
			}
			#endregion Lee Iva

			#region Verifica si aplica revision
			decimal dCant = (decimal) e.Cell.Row.Cells["Cantidad"].Value;
			string stVerif = "Exec DetCompraVerifNegValida " + stIdArt + ", " + dCant.ToString("0.000000", us) + ", " + this.txtIdTipoFactura.Value.ToString();
			if(!Funcion.bEjecutaSQL(cdsSeteoF, stVerif)) return true;
			#endregion verifica si aplica revision

			#region Verifica Negativos o alertas
			string stBodega = cmbBodega.Value.ToString();

			if ((bool) drSeteoF["VerBodega"]) stBodega = e.Cell.Row.Cells["Bodega"].Value.ToString();
			string stidDetC = e.Cell.Row.Cells["idDetCompra"].Value.ToString();
			stVerif = "Exec DetCompraVerifNegativo " + stIdArt + ", " + dCant.ToString("0.000000", us) + ", " + stBodega + ", " + stidDetC + ", " + (int)this.txtIdTipoFactura.Value + ", 0";
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stVerif, true);
			#endregion Verifica Negativos o alertas

			#region Mensaje de Negativo (empieza con -) o Alerta
			if (stMensaje.Length > 0)
			{
				if (stMensaje.StartsWith("-")) // En negativo cambia cantidad a 0
					e.Cell.Row.Cells["Cantidad"].Value = 0;
				MessageBox.Show(stMensaje.Substring(1), "Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return false;
			}
			#endregion Mensaje de Negativo (empieza con -)o Alerta
			return true;
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
		#endregion Grid

		#region Asiento
		private void AsientoAbrir()
		{
			if (txtNumIdAsiento.Value == DBNull.Value || (int) this.txtNumIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe un asiento creado de esta transacción");
				return;
			}

			int IdAsiento = (int) txtNumIdAsiento.Value;
						
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}
		#endregion Asiento
		#endregion Funciones

		decimal dPorcIva = 0.00m;	
		DateTime dtFechaAFin = DateTime.Today;
		DateTime dtFechaAIni = DateTime.Today;
		int idBloqueaTransacciones = 0;
		DateTime dtFechaOrigen = DateTime.Today;

		private void frmDevolucionCompra_Load(object sender, System.EventArgs e)
		{			
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;

			Seguridades();

			Configuracion();

			#region Auditoria
			int iLong = "DEVOLUCION COMPRA".Length;
			if (iLong > 20) iLong = 20;
			string stAudita = "Exec AuditaCrear 53, 6, '" + "DEVOLUCION COMPRA".Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			#endregion Auditoria

			
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Bodega In (1, 20,95)");

			this.cmbProyecto1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ProyectoCarga 6");

			ModificaGrilla(false);
						
			this.txtIdTipoFactura.Value = 6;
			this.txtIdComprobante.Value = 4;
			this.txtNumIdProyecto.Value = 19;
			this.txtNumContadoCredito.Value = 2;
			bAsientoAutomatico = Funcion.bEjecutaSQL(cdsSeteoF, "Exec AsientoAutomaticoVerifica 6");

			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;

			dtFechaAFin = (DateTime)drSeteoC["FechaFin"];
			dtFechaAIni = (DateTime)drSeteoC["FechaIni"];
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			if (idCompra > 0) this.Consulta(idCompra);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtSerieFact_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtNumero.Text.ToString().Length == 0 && (bool)this.chkFiscal.Checked)
				{
					MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;				
				}
				else if ((bool)this.chkFiscal.Checked)
				{
					char Pad = '0';
					this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);
			
					string sSQL = "";
					if (bNuevo) sSQL = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 6 And idCliente = {0} And SerieFactura = '{1}' And Numero = '{2}'", (int)this.txtNumIdCliente.Value, this.txtSerieFact.Value.ToString(), this.txtNumero.Text.ToString());
					else sSQL = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 6 And idCliente = {0} And SerieFactura = '{1}' And Numero = '{2}' And idCompra <> {3}", (int)this.txtNumIdCliente.Value, this.txtSerieFact.Value.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value);
			
					if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) > 0)
					{
						MessageBox.Show("El número ya existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumero.Focus();
						e.Cancel = true;					
					}			
				}
			}
		}

		private void txtSerieFact_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtSerieFact.Text.ToString().Length == 0 && (int)this.txtIdTipoRuc.Value != 3 && (bool)this.chkFiscal.Checked)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					e.Cancel = true;				
				}
				else if (this.txtSerieFact.Text.Length < 6 && (int)this.txtIdTipoRuc.Value != 3 && (bool)this.chkFiscal.Checked)
				{
					MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					e.Cancel = true;				
				}
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{
				if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3) this.txtAutFactura.Focus();
				else
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Focus();
					else 
					{
						this.ultraGrid1.Focus();
						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
					}
				}
			}
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			char Pad = '0';
			this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);
		}

		private void txtAutFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{
				if (miAcceso.CambiarFecha) this.dtFecha.Focus();
				else
				{
					this.ultraGrid1.Focus();
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;				
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Notas de Credito en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void btCrearAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			AsientoAbrir();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, 0, (int)this.txtIdTipoRuc.Value, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
//						this.txtNumIdCliente.Value = miBusqueda.idCliente;
//						this.txtNombre.Text = miBusqueda.Nombre;
						this.txtRuc.Text = miBusqueda.Ruc;
//						this.txtIdTipoRuc.Value = miBusqueda.idTipoRuc;

						if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.txtNumero.Focus();
						else this.txtSerieFact.Focus();	
					}
				}
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			#region Busca Cliente - Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (drC.HasRows)
			{
				this.txtNumIdCliente.Value = (int) drC.GetValue(1);
				this.txtNombre.Text = (string) drC.GetValue(3);
				this.txtRuc.Text = (string) drC.GetValue(2);
				this.txtIdTipoRuc.Value = (int) drC.GetValue(0);
				this.txtIdGrupoCliente.Value = (int) drC.GetValue(4);

				this.chkFiscal.Enabled = false;
			}
			drC.Close();
			#endregion Busca Cliente - Proveedor

			#region S.R.I.
			if (!this.chkFiscal.Checked)
			{
				this.txtSerieFact.Enabled = false;	
				this.txtAutFactura.Enabled = false;	
					
				this.txtSerieFact.Text = "";	
				this.txtAutFactura.Text = "";	
				this.txtNumero.Focus();
			}
			else
			{
				this.txtSerieFact.Enabled = true;	
				this.txtAutFactura.Enabled = true;	
				this.txtSerieFact.Focus();
			}
			#endregion S.R.I.
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//				string sSQL = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.txtIdTipoRuc.Value);
			//				string SMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			//				
			//				if (SMensaje.StartsWith("Error"))
			//				{
			//					MessageBox.Show(SMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtRuc.Focus();
			//					e.Cancel = true;
			//					return;
			//				}
			//
			if (bEdicion)
			{
				if (this.txtRuc.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de identificación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;					
				}
				else 
				{
					string sSQLB = string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString());
					int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLB);
					if (iCount == 0)
					{
						MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();
						e.Cancel = true;						
					}
				}
			}
		}

		private void ActualizaIVA()
		{
			if (this.ultraGrid1.Rows.Count > 0)
			{
				if (this.dtFecha.Value != System.DBNull.Value)
				{
					int iFiscal = 0;
					if (this.chkFiscal.Checked) iFiscal = 1;
					
					if ((int)this.txtNumidOrigen.Value > 0)
					{						
						dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
							dtFechaOrigen.ToString("yyyyMMdd"), iFiscal));
					}
					else 
					{
						dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iFiscal));
					}
				}
			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					string sSQL = string.Format("Select Impuesto From DetCompra Where idArticulo = {0} And idDetCompra = {1}", 
						(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idOrigen"].Value);
					decimal dImpuesto = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

					if (dImpuesto != 0) dr.Cells["Impuesto"].Value = dPorcIva; 
					else dr.Cells["Impuesto"].Value = 0.00m;
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Lista desplegable de Articulos
			if (drSeteoF["VerDesplegable"] != DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				e.Layout.Bands[0].Columns["Articulo"].Hidden = false;
				e.Layout.Bands[0].Columns["Codigo"].Hidden = false;
			
				e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;
			}
			#endregion Lista desplegable de Articulo
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
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
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
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].MaskInput = stInput;
			#endregion
			#region Decimales para el precio
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;
			#endregion
			#region Decimales para el Impuesto
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;
			#endregion
			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			#endregion
		}

		private void Imprimir(int iTipo)
		{			
			#region Seguridad para imprimir y reimprimir y control de vacios
			DateTime dtIni = DateTime.Now;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.ReImprimir && (bool) bImpreso)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			// Cuando no hay registros
			
			#endregion Seguridad para imprimir y reimprimir
			
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + this.txtNumIdCompra.Value;
			if (bAutoImpresor) stReporte += ", 1";
			else stReporte += ", 0";
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			string stFiltro = "{Compra.idCompra} = " + this.txtNumIdCompra.Value.ToString();
			#endregion Nombre de Reporte

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				#region Reporte con procedimiento almacenada
				string stProc = "Exec SerieExisteTipo " + this.txtNumIdCompra.Value.ToString() + ", 'FPVTISP', " + " 4"; //cmbComprobante1.Value.ToString();
				bool bProc = Funcion.bEjecutaSQL(cdsSeteoF, stProc);
				if (bProc == true) 
				{
					stFiltro = "";
					stReporte = stReporte.Substring(0, stReporte.Length - 4) + "SP.rpt";

					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@idCompra";
					discreteVal.Value = this.txtNumIdCompra.Value.ToString();
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					miReporte.crVista.ParameterFieldInfo = paramFields;
				}
				#endregion Reporte con procedimiento almacenada
			
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
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
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

			#region Verifica Cambio
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
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, 2, 0", stCod, (int)this.txtIdTipoFactura.Value, (int) this.cmbBodega.Value);
				if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				#region Mas de una coincidencia o ninguna muestra la lista
				if (iCuenta != 1)
				{ 
					if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));

					int Bodega = (int) this.cmbBodega.Value;

					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, Bodega, (int)this.txtIdTipoFactura.Value))
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

				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				if ((int)this.txtNumidOrigen.Value > 0) dtFecha = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Select Fecha From Compra Where idCompra = {0}", (int)this.txtNumidOrigen.Value));				

				stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 6, '{2}', {3}", IdArt, (int)this.txtIdGrupoCliente.Value, dtFecha.ToString("yyyyMMdd"), (bool)this.chkFiscal.Checked);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
				dr.Read();
				e.Cell.Row.Cells["Codigo"].Value = dr.GetString(0);
				e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
				e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
				e.Cell.Row.Cells["Cantidad"].Value = 1;
				e.Cell.Row.Cells["Precio"].Value = dr.GetDecimal(3);
				e.Cell.Row.Cells["Impuesto"].Value = dr.GetDecimal(4);
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				dr.Close();

				cmdTotal();
        
				VerificaCambio = 2; // Verificaciones y Cambio de precio
				#endregion Asignaciones
			}
			#endregion Codigo o articulo
			#region idArticulo
			if (e.Cell.Column.ToString() == "idArticulo")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value != DBNull.Value)
				{
					int IdArt = (int) e.Cell.Row.Cells["idArticulo"].Value;
					string stSelectU = "Select idUnidad From Articulo Where idArticulo = " + IdArt.ToString();
					e.Cell.Row.Cells["idUnidad"].Value = Funcion.iEscalarSQL(cdsSeteoF, stSelectU, true);
					VerificaCambio = 2;
				}
			}
			#endregion idArticulo
			#region idUnidad o Unidades
			if (e.Cell.Column.ToString() == "idUnidad" || e.Cell.Column.ToString() == "Unidades")
			{
				if (e.Cell.OriginalValue == e.Cell.Value) return;
				CambiaCantidad(e);
				VerificaCambio = 1; // Verificaciones y cambio de precio conficional
			}
			#endregion idUnidad o Unidades

			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["Cantidad"].Value > 0)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Ingrese un Producto antes de Ingresar Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				VerificaCambio = 1; // Verificaciones y cambio de precio conficional
				cmdTotal();
			}
			#endregion Cantidad
			#region Precio
			if (e.Cell.Column.ToString() == "Precio")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["Precio"].Value > 0)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Ingrese un Producto antes de Ingresar el Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				VerificaCambio = 1; // Verificaciones y cambio de precio conficional
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
				VerificaCambio = 1;
			}
			#endregion IVA
			#endregion Fin de Verificacion de Cambio

			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				cmdTotal();
			}
			if (e.Cell.Column.ToString() == "DescuentoArt")
			{				
				cmdTotal();
			}
			#endregion DescuentoPorc o DescuentoArt

			bActualiza = true;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetCompra"].Value = 0;
				e.Row.Cells["idOrigen"].Value = 0;
				e.Row.Cells["DescuentoArt"].Value = 0.00;
				e.Row.Cells["DescuentoPorc"].Value = 0.00;
				e.Row.Cells["Bodega"].Value = (int) this.cmbBodega.Value;
				e.Row.Cells["Signo"].Value = IdSigno;
				e.Row.Cells["Tipo"].Value = 0;				
			}
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			int iTotalLineas = ultraGrid1.Rows.Count;
			string stExec = string.Format("Exec CompraNumeroLineas {0}, {1}", (int)this.txtIdTipoFactura.Value, iTotalLineas);
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);

			if (stMensaje.Length > 0) 
			{
				e.Cancel = true;
				MessageBox.Show(stMensaje); // "No puede crear más lineas");
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
					for (int i=0; i<e.Rows.Length; i++)
					{
						string stIdDetCompra = e.Rows[i].Cells["idDetCompra"].Value.ToString();
						string stExec = "Exec DetCompraBorrarL " + stIdDetCompra;
						string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec, true);
						if (stMensaje.Length > 0)
						{
							MessageBox.Show(stMensaje, "Información");
							e.Cancel = true;
							return;
						}
					}
					e.DisplayPromptMsg = false;
				}
			}
			else
				e.DisplayPromptMsg = false;
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

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			
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
						string sArticulo = ""; 
						sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						int IdArt = 0;
						#region Busqueda servicios
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion)
						{							
							using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(sArticulo, 6))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									IdArt = miBusqueda.idArticulo;
									ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
								}
							}
						}										
						#endregion Busqueda servicios
						#endregion Inventario
						break;
						#endregion F3
					#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
					#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

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
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
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
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al cargar información del artísculo");
			}
		}

		private void frmDevolucionCompra_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir cuando esta editando un registro.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				e.Cancel = true;
			}
		}

		private void frmDevolucionCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if (bImpreso) 
																 {
																	 if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																 }
																 else if (miAcceso.ReImprimir) if (this.btnImprimir.Enabled) this.btnImprimir.Enabled = true;
				if (e.KeyCode == Keys.Escape) this.Close();
				if (e.KeyCode == Keys.F5)
				{
					miAcceso = new Acceso(cdsSeteoF, "0612");
					this.Consulta((int)this.txtNumIdCompra.Value);
				}
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerieFact_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.txtIdTipoRuc.Value != 3) Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) if (this.chkFiscal.Checked) this.txtSerieFact.Focus(); else this.dtFecha.Focus();
		}

		private void txtAutFactura_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtAutFactura.Text.Length == 0 && (int)this.txtIdTipoRuc.Value != 3 && (bool)this.chkFiscal.Checked)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtAutFactura.Focus();
					e.Cancel = true;					
				}
				else if ((this.txtAutFactura.Text.Length != 10 && this.txtAutFactura.Text.Length != 37 && this.txtAutFactura.Text.Length != 49) && (int)this.txtIdTipoRuc.Value != 3 && (bool)this.chkFiscal.Checked)
				{
					MessageBox.Show("Longitud de autorización debe tener 10, 37 o 49 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtAutFactura.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void frmDevolucionCompra_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void txtAutFactura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

			this.lblEstado.Text = "";
			this.txtNumIdCompra.Value = 0;
			this.txtNumidOrigen.Value = 0;
			this.txtNumIdFormaPago.Value = 0;
			this.txtNumIdCliente.Value = 0;
			this.txtNumIdAsiento.Value = 0;
			this.txtNumEstado.Value = 0;
			this.txtNumIdFormaPago.Value = 0;
			this.txtNumIdproyectoZoom.Value = 0;
			this.txtIdCredTributario.Value = 0;
			this.txtIdTipoFormaPago.Value = 0;

			this.txtBuscar.Text = "";
			this.txtBuscar.Enabled = false;

			this.txtIdTipoRuc.Value = 0;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtSerieFact.Text = "";
			this.txtNumero.Text = "";
			this.txtAutFactura.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNotas.Text = "";
			this.txtComprobante.Text = "";
			this.chkFiscal.Checked = false;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.cmbBodega.Enabled = false;
			this.chkFiscal.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtSerieFact.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtAutFactura.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNotas.Enabled = false;
			this.txtComprobante.Enabled = false;
						
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			ModificaGrilla(false);

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleCompraDevolucion 0");

			#region Botones
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAsiento.Enabled = false;
			#endregion Botones

			bAnulado = false;
			bImpreso = false;
			idBloqueaTransacciones = 0;
			dPorcIva = 0.00m;
			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddMonths(-12);
			this.btnNuevo.Focus();
		}

		private void btCrearAsiento_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			this.Consulta((int)this.txtNumIdCompra.Value);

			if (bAnulado)
			{				
				MessageBox.Show("El registro esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (!miAcceso.Anular)
			{				
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				string sMensaje = "ANULAR";
				if (bAnulado) sMensaje = "ELIMINAR";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede '{1}' esta Nota de Crédito, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper(), sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}			
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From HistorialAutorizacionDePago Where idAsiento = {0} And Estado = 0", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("No puede Anular Esta Nota de Credito ya ha sido Utilizada en Pagos de Facturas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
//			else if ((int)this.txtNumEstado.Value == 8)
//			{
//				MessageBox.Show("No Puede Anular Notas de Crédito Liquidadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
//			}
			else
			{
				int idTipoGrupo = 0;
				int iCount = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					string sSQL1 = string.Format("Select idTipoGrupo From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value);
					idTipoGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQL1);
					if (idTipoGrupo == 1 || idTipoGrupo == 6) iCount++; 
				}	

				if (iCount > 0 && (int)this.txtNumEstado.Value == 7)
				{
					MessageBox.Show("No puede editar ya se registro las Seriales de los Articulos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				}
				else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))							
				{
					if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Anular?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulacionCompraVenta {0}, 6, ''", (int)this.txtNumIdCompra.Value), true);
						
						this.Consulta((int)this.txtNumIdCompra.Value);						
					}
				}
			}
		}

		private void Consulta(int idCompra)
		{
			try
			{	
				#region Botones
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAsiento.Enabled = false;
				#endregion Botones
	
				string sSQL = string.Format("Exec ConsultaNCPIndividual {0}", idCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{		
					this.txtNumIdCompra.Value = dr.GetInt32(0);
					if (dr.GetValue(1) != System.DBNull.Value) this.txtNumidOrigen.Value = (int)dr.GetInt32(1);
					this.txtNumEstado.Value = dr.GetInt32(2);
					bImpreso = dr.GetBoolean(3);
					this.cmbBodega.Value = dr.GetInt32(4);
					this.txtNumIdAsiento.Value = dr.GetInt32(5);
					if (dr.GetValue(6) != System.DBNull.Value) this.txtNumIdProyecto.Value = dr.GetInt32(6);
					if (dr.GetValue(7) != System.DBNull.Value) this.txtNumIdproyectoZoom.Value = dr.GetInt32(7);
					this.txtNumContadoCredito.Value = dr.GetInt32(8);
					this.txtNumIdFormaPago.Value = dr.GetInt32(9);
					this.txtNumIdCliente.Value = dr.GetInt32(10);
					this.txtIdTipoRuc.Value = dr.GetInt32(11);
					if (dr.GetValue(12) != System.DBNull.Value) this.txtIdGrupoCliente.Value = dr.GetInt32(12);
					this.txtRuc.Text = dr.GetString(13);
					this.txtNombre.Text = dr.GetString(14);
					if (dr.GetValue(15) != System.DBNull.Value) this.txtSerieFact.Text = dr.GetString(15);
					this.txtNumero.Text = dr.GetString(16);
					if (dr.GetValue(17) != System.DBNull.Value) this.txtAutFactura.Text = dr.GetString(17);
					this.dtFecha.Value = dr.GetDateTime(18);
					if (dr.GetValue(19) != System.DBNull.Value) this.txtComprobante.Text = dr.GetString(19);
					if (dr.GetValue(20) != System.DBNull.Value) this.txtNotas.Text = dr.GetString(20);
					this.txtIva0.Value = dr.GetDecimal(21);
					this.txtIva.Value = dr.GetDecimal(22);
					this.txtDescIva0Total.Value = dr.GetDecimal(23);
					this.txtDescIvaTotal.Value = dr.GetDecimal(24);
					this.txtIvaTotal.Value = dr.GetDecimal(25);
					this.txtTotal.Value = dr.GetDecimal(26);
					bAnulado = dr.GetBoolean(27);
					if (dr.GetValue(28) != System.DBNull.Value) this.txtIdCredTributario.Value = dr.GetInt32(28);					
					this.chkFiscal.Checked = dr.GetBoolean(29);
					this.lblEstado.Text = dr.GetString(30);
					this.lblUsuario.Text = dr.GetString(31);
				}
				dr.Close();

				string sSQLDetalle = string.Format("Exec ConsultaDetalleCompraDevolucion {0}", idCompra);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLDetalle);
					
				#region Botones
				if (!bAnulado)
				{
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
					{
						if (miAcceso.Editar) //if ((int)this.txtNumEstado.Value < 8) 
																	 this.btnEditar.Enabled = true;          
						if (miAcceso.Anular) this.btnAnular.Enabled = true;					
					}

					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
				}				
				
				this.btnCancelar.Enabled = true;
				#endregion Botones				          

				this.txtBuscar.Enabled = false;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaDevCompra Busqueda = new frmBuscaDevCompra())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Editar esta Nota De Crédito, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}
			else 
			{	
				int iCount = 0;

				if ((int)this.txtNumEstado.Value == 7)
				{
					int idTipoGrupo = 0;
				
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						string sSQL1 = string.Format("Select idTipoGrupo From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value);
						idTipoGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQL1);
						if (idTipoGrupo == 1 || idTipoGrupo == 6) iCount++; 
					}	
				}

				if (iCount > 0)
				{
					MessageBox.Show("No puede editar ya se registro las Seriales de los Articulos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				}	
				else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From HistorialAutorizacionDePago Where idAsiento = {0} And Estado = 1", (int)this.txtNumIdCompra.Value)) > 0)
				{
					MessageBox.Show("No puede editar Esta Nota de Credito ya ha sido Utilizada en Pagos de Facturas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
//				else if ((int)this.txtNumEstado.Value == 8)
//				{				
//					MessageBox.Show("No puede editar una Nota de Credito Liquidada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
//				}	
				else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))			
				{
					if (this.chkFiscal.Checked)
					{
						this.txtSerieFact.Enabled = true;
						this.txtAutFactura.Enabled = true;
					}

					this.txtNumero.Enabled = true;
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					this.txtNotas.Enabled = true;
			
					this.ModificaGrilla (true);

					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnGuardar.Enabled = true;
					this.btnCancelar.Enabled = true;
					this.btnAnular.Enabled = false;
					this.btnAsiento.Enabled = false;
			
					if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3) this.txtSerieFact.Focus();
					else this.txtNumero.Focus();

					string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
					idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

					bNuevo = false;
					bEdicion = true;
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			this.txtNumEstado.Value = 1;
			this.txtIdTipoRuc.Value = 1;
			this.txtNumEstado.Value = 1;
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			this.txtNumIdFormaPago.Value = 6;
			this.cmbBodega.Value = 1;
			this.txtIdTipoFormaPago.Value = 2;

			this.chkFiscal.Checked = true;
			this.chkFiscal.Enabled = true;

			this.txtRuc.Enabled = true;
			this.txtSerieFact.Enabled = true;
			this.txtNumero.Enabled = true;
			this.txtAutFactura.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			this.cmbBodega.Enabled = true;
			this.txtNotas.Enabled = true;

			ModificaGrilla(true);

			this.btnConsultar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnGuardar.Enabled = true;

			this.txtBuscar.Enabled = true;			
			this.txtBuscar.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			int iTipo = 0;
						
			Cursor = Cursors.WaitCursor;			
			string stFiltro = "{Compra.idCompra} = " + this.txtNumIdCompra.Value.ToString();
					
			#region Imprime el reporte
			using (Reporte miReporte = new Reporte("DevolucionCompra.Rpt", stFiltro))
			{
				#region Reporte con procedimiento almacenada
				string stProc = "Exec SerieExisteTipo " + this.txtNumIdCompra.Value.ToString() + ", 'FPVTISP', " + " 15 ";
				bool bProc = Funcion.bEjecutaSQL(cdsSeteoF, stProc);
				if (bProc == true) 
				{
					stFiltro = "";
					string stReporte = "DevolucionCompraSP.rpt";

					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@idCompra";
					discreteVal.Value = this.txtNumIdCompra.Value.ToString();
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					miReporte.crVista.ParameterFieldInfo = paramFields;
				}
				#endregion Reporte con procedimiento almacenada
			
				if (iTipo == 0) // Muestra en pantalla
					miReporte.ShowDialog();
				if (iTipo >= 1) // Imprime una o mas copias
				{
					string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
					string stImpresora = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraReporteImpresora 'DevolucionCompra.Rpt'");
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

			#region Marca como impreso
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + this.txtNumIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			bImpreso = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso
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
			#region Verifica Cero
			if ((decimal) txtTotal.Value == 0) 
			{
				MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Verifica Cero			
			#region Encabezado
			#region RUC CEDULA PASAPORTE
			if (this.txtRuc.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de identificación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			string sSQLVI = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.txtIdTipoRuc.Value);
			string SMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVI);
				
			if (SMensaje.StartsWith("Error"))
			{
				MessageBox.Show(SMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			
			string sSQLB = string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString());
			int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLB);
			if (iCount == 0)
			{
				MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			#endregion RUC CEDULA PASAPORTE
			
			if (this.chkFiscal.Checked)
			{
				#region Serie
				if (this.txtSerieFact.Text.ToString().Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}

				if (this.txtSerieFact.Text.ToString().Length < 6 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}
				#endregion Serie

				#region Numero
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			
				string sSQL1 = "";
				if (bNuevo) sSQL1 = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 6 And idCliente = {0} And SerieFactura = '{1}' And Numero = '{2}'", (int)this.txtNumIdCliente.Value, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString());
				else sSQL1 = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 6 And idCliente = {0} And SerieFactura = '{1}' And Numero = '{2}' And idCompra <> {3}", (int)this.txtNumIdCliente.Value, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value);
			
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL1) > 0)
				{
					MessageBox.Show("El número ya existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}	
				#endregion Numero

				#region Autorización
				if (this.txtAutFactura.Text.Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtAutFactura.Focus();
					return;
				}
				if ((this.txtAutFactura.Text.Length != 10 && this.txtAutFactura.Text.Length != 37 && this.txtAutFactura.Text.Length != 49) && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("La longitud de autorización debe tener 10, 37 o 49 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtAutFactura.Focus();
					return;
				}
				#endregion Autorización
			}
			else if (bNuevo)
			{
				string sSQLNumero = string.Format("Select Count(*) From Compra Where idTipoFactura = 6 And Fiscal = 0");
				this.txtNumero.Text = "NC-" + Funcion.iEscalarSQL(cdsSeteoF, sSQLNumero).ToString();
			}

			#region Verificacion de Fecha
			if (this.dtFecha.Value == System.DBNull.Value) 
			{
				MessageBox.Show("Ingrese la fecha de la Nota de Crédito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}	
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Notas de Credito en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Verificacion de Fecha		
			#endregion Encabezado											
			#region Detalle
			#region Cantidad Precio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value || dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
					return;
				}

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int IdArt = (int) dr.Cells["idArticulo"].Value;

					decimal dPrecio = (decimal) dr.Cells["Precio"].Value;
					
					if (dPrecio == 0.00m && dr.Cells["idArticulo"].Value != System.DBNull.Value) 
					{
						MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Cantidad Precio		
			#region Verifica Ingreso en el grid
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el detalle de la Nota de Crédito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}
			#endregion Verifica Ingreso en el grid
			#endregion Detalle

			#region Valida Existencia
			if (!Validacion.bValidaExistencias(this.ultraGrid1, (DateTime)this.dtFecha.Value, (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, 6, 0, cdsSeteoF)) return;
			#endregion Valida Existencia
			#endregion Validaciones

			this.cmdTotal();

			#region Mensaje
			//			if (bNuevo && (int)this.txtNumidOrigen.Value > 0) 
			//				if (DialogResult.No == MessageBox.Show("¿Esta seguro de guardar, esta accion modificara los pagos de la Factura N. " + this.txtComprobante.Text.ToString() + "?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			#endregion Mensaje

			#region Valida Estado por tipo de articulo I / B-S
			//			int idTipoGrupo = 0;
			//
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			//			{
			//				string sSQL1 = string.Format("Select idTipoGrupo From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value);
			//				idTipoGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQL1);
			//				if (idTipoGrupo == 1) iCount++; 
			//				if (idTipoGrupo == 6) iCount++; 
			//			}	
			//
			//			if (iCount == 0) this.txtNumEstado.Value = 7;
			//			else 
			
			this.txtNumEstado.Value = 7;

			int idTipoPago = 1;
			if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) idTipoPago = 2;
			#endregion Valida Estado por tipo de articulo I / S

			#region Guardar
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

					#region Encabezado
					DateTime dtFecha = (DateTime)this.dtFecha.Value;

					string sSQL = string.Format("Exec GuardaEncDevCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', '{9}', {10}, {11}, {12}, {13}, {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}", 
						(int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtIdSucursal.Value, 
						(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, (int)this.txtNumEstado.Value, 
						(int)this.txtIdTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), 
						(int)this.txtNumContadoCredito.Value, (int)this.txtNumIdFormaPago.Value, 
						(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value, (int)this.txtNumIdproyectoZoom.Value, (int)this.txtNumIdAsiento.Value,
						this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), this.txtAutFactura.Text.ToString(), this.txtComprobante.Text.ToString(), 
						dtFecha.ToString("yyyyMMdd"), this.txtNotas.Text.ToString(), 
						(decimal)this.txtIva0.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
						(decimal)this.txtTotal.Value, (int)this.txtIdCredTributario.Value, idTipoPago, (int)this.txtIdTipoFormaPago.Value, this.chkFiscal.Checked);
					oCmdActualiza.CommandText = sSQL;
					this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Encabezado

					#region Detalle
					int idProyecto = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if (dr.Cells["idProyecto"].Value != System.DBNull.Value) idProyecto = (int)dr.Cells["idProyecto"].Value;
						string sSQLDetalle = string.Format("Exec GuardaDetDevCompra {0}, {1}, {2}, {3}, '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}", 
							(int)dr.Cells["idDetCompra"].Value, (int)this.txtNumIdCompra.Value, (int)dr.Cells["idUnidad"].Value, 
							(int)dr.Cells["idArticulo"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value.ToString(), 
							(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, (decimal)dr.Cells["DescuentoArt"].Value,
							(decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["SubTotal"].Value, 
							(int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["idOrigen"].Value, (int)dr.Cells["Tipo"].Value, idProyecto);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}

					string sSQLActDetalle = string.Format("Exec ConsultaDetalleCompraDevolucion {0}", (int)this.txtNumIdCompra.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);						
					this.ultraGrid1.DataSource = oDTDetalle;						
					#endregion Detalle
					
					#region Asiento Automatico
					if (bAsientoAutomatico) 
					{				
						#region Crea Asiento
						string sSQLAsiento = string.Format("Exec CreaAsiDevCompra {0}", this.txtNumIdCompra.Value.ToString());
						oCmdActualiza.CommandText = sSQLAsiento;
						oCmdActualiza.ExecuteNonQuery();
				
						oCmdActualiza.CommandText = string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
						this.txtNumIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

						Cursor = Cursors.Default;
						#endregion Crea Asiento
					}
					#endregion Asiento Automatico

					#region Saldo 
					string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 6, 6", (int)this.txtNumIdCompra.Value);
					oCmdActualiza.CommandText = sSQLSaldo;
					oCmdActualiza.ExecuteNonQuery();					
					#endregion Saldo 

					if (!bNuevo)
					{
						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
						oCmdActualiza.CommandText = sSQLBloqueo;
						oCmdActualiza.ExecuteNonQuery();
					}

					oTransaction.Commit();	

					#region Controles y estado
					#region Botones
					this.txtBuscar.Enabled = false;
					this.txtRuc.Enabled = false;
					this.txtSerieFact.Enabled = false;
					this.txtNumero.Enabled = false;
					this.txtAutFactura.Enabled = false;
					this.dtFecha.Enabled = false;
					this.cmbBodega.Enabled = false;
					ModificaGrilla(false);
					this.txtNotas.Enabled = false;
					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular)	this.btnAnular.Enabled = true;
					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;
					#endregion Botones

					this.chkFiscal.Enabled = false;

					bNuevo = false;
					bEdicion = false;

					if ((int)this.txtNumEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
					else if ((int)this.txtNumEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
					#endregion Controles y estado

					bEliminaAlValidar = false;
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
			#endregion Guardar
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (frmBuscaFacturasCruzaNC miBusqueda = new frmBuscaFacturasCruzaNC(4))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						try
						{	
							#region Maestro
							this.txtNumIdCompra.Value = 0;
							this.txtNumidOrigen.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
							//bConsignacion = (bool)miBusqueda.grdFacturas.ActiveRow.Cells["Consignacion"].Value;
							this.dtFechaOrigen = Convert.ToDateTime(miBusqueda.grdFacturas.ActiveRow.Cells["Fecha"].Value);
							this.txtNumIdFormaPago.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idFormaPago"].Value;
							this.txtIdCredTributario.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCredTributario"].Value;
							this.txtComprobante.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Numero"].Value.ToString();
							this.txtIdTipoRuc.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idTipoRuc"].Value;
							this.txtNumIdCliente.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCliente"].Value;
							this.txtRuc.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Ruc"].Value.ToString();
							this.txtNombre.Text = miBusqueda.grdFacturas.ActiveRow.Cells["Nombre"].Value.ToString();
							this.txtIva0.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["SubtotalExcento"].Value;
							this.txtDescIva0Total.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Descuento0"].Value;
							this.txtIva.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["SubtotalIva"].Value;
							this.txtDescIvaTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Descuento"].Value;
							this.txtIvaTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Iva"].Value;
							this.txtTotal.Value = (decimal)miBusqueda.grdFacturas.ActiveRow.Cells["Total"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["idGrupoCliente"].Value != System.DBNull.Value) this.txtIdGrupoCliente.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idGrupoCliente"].Value;
							if (miBusqueda.grdFacturas.ActiveRow.Cells["idProyectoZoom"].Value != System.DBNull.Value) this.txtNumIdproyectoZoom.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idProyectoZoom"].Value; 
							this.txtIdTipoFormaPago.Value = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idTipoFormaPago"].Value;
							#endregion Maestro

							this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCompraNC {0}", (int)this.txtNumidOrigen.Value));
				
							this.txtRuc.Enabled = false;
							
							if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3) 
							{
								this.chkFiscal.Checked = true;
								this.txtSerieFact.Focus();
							}
							else
							{
								this.chkFiscal.Checked = false;
								this.txtSerieFact.Enabled = false;
								this.txtAutFactura.Enabled = false;
								this.txtNumero.Focus();
							}
							this.chkFiscal.Enabled = false;

							this.dtFecha.CalendarInfo.MinDate = this.dtFechaOrigen;
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnGuardar.Focus();
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);

			DateTime dtFechaAsi = (DateTime)dtFecha.Value;

			if (bAnulado) 
			{
				MessageBox.Show("No puede crear un asiento de una transacción Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				MessageBox.Show(string.Format("No Puede Crear Asiento de esta Nota de Crédito, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFechaAsi.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}	
			else if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}			
			else
			{	
				string stExisteAsiento = "Exec AsientoMaestroExiste "	+ this.txtNumIdCompra.Value.ToString();
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					Cursor = Cursors.Default;					
				}
				else
				{
					stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 2", true);
					if (stMensaje.Length > 0)
					{
						if (DialogResult.Yes == MessageBox.Show(stMensaje, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 3", true);

							string stGeneraCompra = string.Format("Exec CreaAsiDevCompra {0}", (int) this.txtNumIdCompra.Value);
							Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);

							MessageBox.Show("Asiento Generado con Exito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));

							Cursor = Cursors.Default;
						}
					}
					else if (DialogResult.Yes == MessageBox.Show("¿Desea Crear Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{						
						string stGeneraCompra = string.Format("Exec CreaAsiDevCompra {0}", (int) this.txtNumIdCompra.Value);
						Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);

						MessageBox.Show("Asiento Generado con Exito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value));

						Cursor = Cursors.Default;
					}
				}				
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			cmdTotal();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.ultraGrid1.Focus();
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.txtRuc.Enabled) this.txtRuc.Focus();
				else this.txtSerieFact.Focus();
			}
		}

		private void chkFiscal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion)return;

			if (this.chkFiscal.Checked)
			{
				this.txtRuc.Enabled = false;
				this.txtSerieFact.Enabled = true;
				this.txtNumero.Enabled = true;
				this.txtAutFactura.Enabled = true;
			}
			else
			{
				this.txtRuc.Enabled = true;
				this.txtSerieFact.Enabled = false;
				this.txtNumero.Enabled = false;
				this.txtAutFactura.Enabled = false;

				this.txtSerieFact.Text = "";
				this.txtNumero.Text = "";
				this.txtAutFactura.Text = "";
			}

			ActualizaIVA();
		}

		private void chkFiscal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
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
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Nota De Credito esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						else if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdAsiento.Value), true))
						{
							MessageBox.Show("El Asiento de esta Nota De Credito fue eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						else //if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 0) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdCompra.Value), true))
						{
							Cursor = Cursors.WaitCursor;
							frmAsientos Asientos = new frmAsientos((int) this.txtNumIdAsiento.Value);
							Asientos.MdiParent = MdiParent;
							Asientos.Show();
							Cursor = Cursors.Default;
						}						
					}
				}
			}			
		}

		private void txtComprobante_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				int iFiscal = 0;
				if (this.chkFiscal.Checked) iFiscal = 1;			

				if ((int)this.txtNumidOrigen.Value > 0)		
				{										
					dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
						dtFechaOrigen.ToString("yyyyMMdd"), iFiscal));
				}
				else 
				{
					dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iFiscal));
				}

				ActualizaIVA();
			}
		}
	}
}

