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
	/// Summary description for frmCompra.
	/// </summary>
	public class frmCompra : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsSeteo;
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
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		private System.Windows.Forms.Label lblImportacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbImportacionNombre;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito3;
		private System.Windows.Forms.Label label69;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito2;
		private System.Windows.Forms.Label label70;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito1;
		private System.Windows.Forms.Label label71;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox gbBotones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdproyectoZoom;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumidOrigen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private System.Windows.Forms.Button btnBuscarOrdenDeCompra;
		private System.Windows.Forms.Button btnRetencion;
		private System.ComponentModel.IContainer components;


		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFormaPago;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbImportacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCaducidad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoDePago;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCreditoTributario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.Windows.Forms.Label lblOrigen;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyectoZOOM;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTiposDeGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito4;
		private System.Windows.Forms.Button btnProforma;

		#region Variables Globales
		decimal dPorcIva = 0.00m;		
		int iOrigen = 0;
		int iEstadoOrigen = 9;
		int IdCompra = 0;
		bool BBuscar = false;
		bool BEditar = false;
		bool bAsientoAutomatico = false;
		bool bImpreso = false;
		bool bEliminaAlValidar = false;
		bool bNuevo = false;
		bool bEdicion = false;
		int IdSigno = 1;
		private Acceso miAcceso;
		private bool bActualiza = true;
		bool bRastreo = false;
		bool bAnulado = false;
		bool bGastosPorProyecto = false;
		int idBloqueaTransacciones = 0;
		bool bProforma = false;
		bool bLiquidacion = false;
		decimal dValorProforma = 0.00m;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaPeriodo;
		private System.Windows.Forms.Label lblFechaPEriodo;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito5;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechadeEmicion;
		private System.Windows.Forms.Button btnSubir;
		private System.Windows.Forms.Button btnDistribucion;
		private System.Windows.Forms.ContextMenu mnuDistribucion;
		private System.Windows.Forms.MenuItem mnuAgencias;
		private System.Windows.Forms.MenuItem mnuAdministrativos;
		private System.Windows.Forms.MenuItem mnuGeneral;
		private System.Windows.Forms.Button btnTemporal;
		private System.Windows.Forms.Button btnDias;
		private System.Windows.Forms.MenuItem mnuAgenciasI;
		private System.Windows.Forms.MenuItem mnuAdministrativosI;
		private System.Windows.Forms.MenuItem mnuGeneralI;		
		CultureInfo us = new CultureInfo("en-US");		
		#endregion Variables Globales

		public frmCompra(int idCompra, bool bBuscar)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IdCompra = idCompra;
			BBuscar = bBuscar;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmCompra(int idCompra, bool bBuscar,bool bEditar)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IdCompra = idCompra;
			BBuscar = bBuscar;
			BEditar = bEditar;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmCompra()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		
		string sCSRUC = "";
		bool bConsignacion = false;
		int iCSidFormaPago = 0;

		public frmCompra(bool BConsignacion, string SCSRUC, int ICSidFormaPago)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			sCSRUC = SCSRUC;
			bConsignacion = BConsignacion;
			iCSidFormaPago = ICSidFormaPago; 
			
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCompra));
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valores");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescNC");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion", -1, "cmbImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto", -1, "cmbTiposDeGastos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Courier");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Arancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EsVentaVerde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompraVV");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Courier");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Arancel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EsVentaVerde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompraVV");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaDePago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCredTributario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CredTributario");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comprobante");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDeCompraCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenDeCompraCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyectoZoom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyectoZoom");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripción");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
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
			this.label1 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblImportacion = new System.Windows.Forms.Label();
			this.cmbImportacionNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.txtDiasCredito3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label69 = new System.Windows.Forms.Label();
			this.txtDiasCredito2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label70 = new System.Windows.Forms.Label();
			this.txtDiasCredito1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label71 = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnBuscarOrdenDeCompra = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNumIdproyectoZoom = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumidOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnRetencion = new System.Windows.Forms.Button();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbImportacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdTipoDePago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbCreditoTributario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblOrigen = new System.Windows.Forms.Label();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnProforma = new System.Windows.Forms.Button();
			this.btnSubir = new System.Windows.Forms.Button();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProyectoZOOM = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTiposDeGastos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.txtDiasCredito4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblFechaPEriodo = new System.Windows.Forms.Label();
			this.dtFechaPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDiasCredito5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.dtFechadeEmicion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnDistribucion = new System.Windows.Forms.Button();
			this.mnuDistribucion = new System.Windows.Forms.ContextMenu();
			this.mnuAgencias = new System.Windows.Forms.MenuItem();
			this.mnuAdministrativos = new System.Windows.Forms.MenuItem();
			this.mnuGeneral = new System.Windows.Forms.MenuItem();
			this.mnuAgenciasI = new System.Windows.Forms.MenuItem();
			this.mnuAdministrativosI = new System.Windows.Forms.MenuItem();
			this.mnuGeneralI = new System.Windows.Forms.MenuItem();
			this.btnTemporal = new System.Windows.Forms.Button();
			this.btnDias = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacionNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoDePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyectoZOOM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechadeEmicion)).BeginInit();
			this.SuspendLayout();
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
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(797, 363);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(45, 16);
			this.lblSubTotal.TabIndex = 164;
			this.lblSubTotal.Text = "Subtotal";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(927, 411);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 162;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(797, 387);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(57, 16);
			this.lblDescuento.TabIndex = 160;
			this.lblDescuento.Text = "Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(1021, 336);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 158;
			this.lblIvaSubt.Text = "IVA";
			this.lblIvaSubt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(925, 336);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 156;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance1;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(877, 432);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(168, 34);
			this.txtTotal.TabIndex = 178;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance2;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(965, 408);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 177;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance3;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(877, 384);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 175;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance4;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(965, 384);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 172;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance5;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(965, 360);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 170;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance6;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(877, 360);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 168;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(797, 441);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 179;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(216, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 182;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance7;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(88, 5);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 1;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.TextChanged += new System.EventHandler(this.txtRuc_TextChanged);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance8;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(272, 5);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(352, 22);
			this.txtNombre.TabIndex = 2;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(71, 16);
			this.label2.TabIndex = 252;
			this.label2.Text = "Identificación";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtComprobante
			// 
			this.txtComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance9;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(560, 333);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(120, 22);
			this.txtComprobante.TabIndex = 265;
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(472, 336);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 264;
			this.lblComprobante.Text = "Comprobante";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(640, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 263;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(640, 38);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 269;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance10;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(688, 36);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 8;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(456, 38);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 274;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance11;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(512, 35);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 7;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtSerieFact
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance12;
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Enabled = false;
			this.txtSerieFact.Location = new System.Drawing.Point(376, 35);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(64, 22);
			this.txtSerieFact.TabIndex = 6;
			this.txtSerieFact.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerieFact_KeyDown);
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			this.txtSerieFact.ValueChanged += new System.EventHandler(this.txtSerieFact_ValueChanged);
			// 
			// chkConsignacion
			// 
			this.chkConsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance13;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(696, 333);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 17;
			this.chkConsignacion.Text = "Consignación";
			this.chkConsignacion.Visible = false;
			this.chkConsignacion.CheckedChanged += new System.EventHandler(this.chkConsignacion_CheckedChanged);
			// 
			// lblImportacion
			// 
			this.lblImportacion.AutoSize = true;
			this.lblImportacion.Location = new System.Drawing.Point(312, 68);
			this.lblImportacion.Name = "lblImportacion";
			this.lblImportacion.Size = new System.Drawing.Size(64, 16);
			this.lblImportacion.TabIndex = 288;
			this.lblImportacion.Text = "Importación";
			this.lblImportacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbImportacionNombre
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbImportacionNombre.Appearance = appearance14;
			this.cmbImportacionNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbImportacionNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbImportacionNombre.DataSource = this.ultraDataSource4;
			ultraGridBand1.AddButtonCaption = "CompraCredTribut";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 296;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbImportacionNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbImportacionNombre.DisplayMember = "Nombre";
			this.cmbImportacionNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbImportacionNombre.Enabled = false;
			this.cmbImportacionNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbImportacionNombre.Location = new System.Drawing.Point(392, 66);
			this.cmbImportacionNombre.Name = "cmbImportacionNombre";
			this.cmbImportacionNombre.Size = new System.Drawing.Size(232, 21);
			this.cmbImportacionNombre.TabIndex = 15;
			this.cmbImportacionNombre.ValueMember = "idImportacionNombre";
			this.cmbImportacionNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbImportacionNombre_KeyDown);
			this.cmbImportacionNombre.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbImportacionNombre_MouseDown);
			this.cmbImportacionNombre.Validating += new System.ComponentModel.CancelEventHandler(this.cmbImportacionNombre_Validating);
			this.cmbImportacionNombre.ValueChanged += new System.EventHandler(this.cmbImportacionNombre_ValueChanged);
			this.cmbImportacionNombre.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbImportacionNombre_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(336, 38);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 289;
			this.label3.Text = "Serie";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(480, 416);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(56, 16);
			this.label11.TabIndex = 293;
			this.label11.Text = "Caducidad";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Visible = false;
			// 
			// txtAutFactura
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance15;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Enabled = false;
			this.txtAutFactura.Location = new System.Drawing.Point(88, 35);
			this.txtAutFactura.MaxLength = 49;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(232, 22);
			this.txtAutFactura.TabIndex = 5;
			this.txtAutFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAutFactura_KeyDown);
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			this.txtAutFactura.Validating += new System.ComponentModel.CancelEventHandler(this.txtAutFactura_Validating);
			this.txtAutFactura.ValueChanged += new System.EventHandler(this.txtAutFactura_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(8, 38);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 290;
			this.label31.Text = "Autorización";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance16;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 65);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(200, 22);
			this.cmbFormaPago.TabIndex = 10;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 68);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 295;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasCredito3
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito3.Appearance = appearance17;
			this.txtDiasCredito3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito3.Enabled = false;
			this.txtDiasCredito3.Location = new System.Drawing.Point(448, 125);
			this.txtDiasCredito3.MinValue = 0;
			this.txtDiasCredito3.Name = "txtDiasCredito3";
			this.txtDiasCredito3.PromptChar = ' ';
			this.txtDiasCredito3.Size = new System.Drawing.Size(64, 22);
			this.txtDiasCredito3.TabIndex = 13;
			this.txtDiasCredito3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito3_KeyDown);
			this.txtDiasCredito3.Click += new System.EventHandler(this.txtDiasCredito3_Click);
			this.txtDiasCredito3.ValueChanged += new System.EventHandler(this.txtDiasCredito3_ValueChanged);
			this.txtDiasCredito3.Enter += new System.EventHandler(this.txtDiasCredito3_Enter);
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.BackColor = System.Drawing.Color.Transparent;
			this.label69.Enabled = false;
			this.label69.Location = new System.Drawing.Point(352, 128);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(78, 16);
			this.label69.TabIndex = 301;
			this.label69.Text = "Días Crédito 3:";
			this.label69.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasCredito2
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito2.Appearance = appearance18;
			this.txtDiasCredito2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito2.Enabled = false;
			this.txtDiasCredito2.Location = new System.Drawing.Point(272, 125);
			this.txtDiasCredito2.MinValue = 0;
			this.txtDiasCredito2.Name = "txtDiasCredito2";
			this.txtDiasCredito2.PromptChar = ' ';
			this.txtDiasCredito2.Size = new System.Drawing.Size(64, 22);
			this.txtDiasCredito2.TabIndex = 12;
			this.txtDiasCredito2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito2_KeyDown);
			this.txtDiasCredito2.Click += new System.EventHandler(this.txtDiasCredito2_Click);
			this.txtDiasCredito2.ValueChanged += new System.EventHandler(this.txtDiasCredito2_ValueChanged);
			this.txtDiasCredito2.Enter += new System.EventHandler(this.txtDiasCredito2_Enter);
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.BackColor = System.Drawing.Color.Transparent;
			this.label70.Enabled = false;
			this.label70.Location = new System.Drawing.Point(184, 128);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(78, 16);
			this.label70.TabIndex = 300;
			this.label70.Text = "Días Crédito 2:";
			this.label70.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasCredito1
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito1.Appearance = appearance19;
			this.txtDiasCredito1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito1.Enabled = false;
			this.txtDiasCredito1.Location = new System.Drawing.Point(96, 125);
			this.txtDiasCredito1.MinValue = 0;
			this.txtDiasCredito1.Name = "txtDiasCredito1";
			this.txtDiasCredito1.PromptChar = ' ';
			this.txtDiasCredito1.Size = new System.Drawing.Size(64, 22);
			this.txtDiasCredito1.TabIndex = 11;
			this.txtDiasCredito1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito1_KeyDown);
			this.txtDiasCredito1.Click += new System.EventHandler(this.txtDiasCredito1_Click);
			this.txtDiasCredito1.ValueChanged += new System.EventHandler(this.txtDiasCredito1_ValueChanged);
			this.txtDiasCredito1.Enter += new System.EventHandler(this.txtDiasCredito1_Enter);
			// 
			// label71
			// 
			this.label71.AutoSize = true;
			this.label71.BackColor = System.Drawing.Color.Transparent;
			this.label71.Enabled = false;
			this.label71.Location = new System.Drawing.Point(8, 128);
			this.label71.Name = "label71";
			this.label71.Size = new System.Drawing.Size(78, 16);
			this.label71.TabIndex = 299;
			this.label71.Text = "Días Crédito 1:";
			this.label71.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbProyecto
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance20;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(859, 6);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(168, 21);
			this.cmbProyecto.TabIndex = 4;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(760, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(85, 16);
			this.label4.TabIndex = 303;
			this.label4.Text = "Centro de Costo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance21.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance21;
			ultraGridBand4.AddButtonCaption = "_DetCompra";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 12;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 73;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn11.Width = 53;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 200;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 250;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance22;
			ultraGridColumn14.Format = "";
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.MaskInput = "";
			ultraGridColumn14.NullText = "0";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 70;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance23;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 68;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance24;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "IVA";
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 50;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance25;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 9;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance26;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Descuento";
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 38;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance27;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "% Desc.";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 65;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance28;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.PromptChar = ' ';
			ultraGridColumn20.Width = 90;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance29;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Desc N C";
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn22.Header.Caption = "Centro de costo";
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn22.Width = 150;
			ultraGridColumn23.Header.Caption = "Importación";
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 150;
			ultraGridColumn24.Header.Caption = "Gasto";
			ultraGridColumn24.Header.VisiblePosition = 18;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn24.Width = 180;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance30;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.MaskInput = "nnnnnnnnn";
			ultraGridColumn25.Width = 14;
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn27.Hidden = true;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance31;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 69;
			ultraGridColumn29.Header.VisiblePosition = 22;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 24;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 25;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 26;
			ultraGridColumn33.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance32;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance33;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 9F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance34;
			appearance35.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance35.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance37;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 160);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1039, 168);
			this.ultraGrid1.TabIndex = 18;
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
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
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
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
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn27.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn28.DataType = typeof(bool);
			ultraDataColumn28.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn29.DataType = typeof(short);
			ultraDataColumn29.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(472, 408);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 305;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Width = 180;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(680, 6);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(64, 21);
			this.cmbBodega.TabIndex = 3;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(402, 488);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 26;
			this.btnAsiento.Text = "&Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnAsiento, "Ver Asiento Con Click Izquierdo Sobreescribe el Asiento/Con Click Derecho Abre el" +
				" Asiento");
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnBuscarOrdenDeCompra
			// 
			this.btnBuscarOrdenDeCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBuscarOrdenDeCompra.CausesValidation = false;
			this.btnBuscarOrdenDeCompra.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscarOrdenDeCompra.ImageIndex = 67;
			this.btnBuscarOrdenDeCompra.ImageList = this.imageList1;
			this.btnBuscarOrdenDeCompra.Location = new System.Drawing.Point(568, 488);
			this.btnBuscarOrdenDeCompra.Name = "btnBuscarOrdenDeCompra";
			this.btnBuscarOrdenDeCompra.Size = new System.Drawing.Size(115, 23);
			this.btnBuscarOrdenDeCompra.TabIndex = 24;
			this.btnBuscarOrdenDeCompra.Text = "&Orden de Compra";
			this.btnBuscarOrdenDeCompra.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnBuscarOrdenDeCompra, "Busca Ordenes de Compra Procesadas");
			this.btnBuscarOrdenDeCompra.Click += new System.EventHandler(this.btnBuscarOrdenDeCompra_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 488);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Crea una Nueva Factura");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 488);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 28;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnConsultar, "Permte Buscar una Factura");
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(164, 488);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 27;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnEditar, "Permite Modificar la Factura Actual");
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 488);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 21;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnImprimir, "Imprime la Factura Actual");
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			this.btnImprimir.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnImprimir_MouseDown);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 488);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 20;
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
			this.btnAnular.Location = new System.Drawing.Point(843, 488);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 309;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnAnular, "Anula una Factura ");
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(922, 488);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 22;
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
			this.btnSalir.Location = new System.Drawing.Point(997, 488);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 23;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 472);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1055, 8);
			this.gbBotones.TabIndex = 318;
			this.gbBotones.TabStop = false;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance38;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 360);
			this.txtNotas.MaxLength = 600;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(448, 104);
			this.txtNotas.TabIndex = 19;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 336);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(33, 16);
			this.label5.TabIndex = 320;
			this.label5.Text = "Notas";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumIdproyectoZoom
			// 
			this.txtNumIdproyectoZoom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdproyectoZoom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdproyectoZoom.Enabled = false;
			this.txtNumIdproyectoZoom.Location = new System.Drawing.Point(655, 440);
			this.txtNumIdproyectoZoom.Name = "txtNumIdproyectoZoom";
			this.txtNumIdproyectoZoom.PromptChar = ' ';
			this.txtNumIdproyectoZoom.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdproyectoZoom.TabIndex = 334;
			this.txtNumIdproyectoZoom.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(623, 440);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 331;
			this.txtNumEstado.Visible = false;
			this.txtNumEstado.ValueChanged += new System.EventHandler(this.txtNumEstado_ValueChanged);
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(607, 440);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 330;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(751, 440);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 329;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(591, 440);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 328;
			this.txtIdSucursal.Visible = false;
			// 
			// txtNumidOrigen1
			// 
			this.txtNumidOrigen1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen1.Enabled = false;
			this.txtNumidOrigen1.Location = new System.Drawing.Point(575, 440);
			this.txtNumidOrigen1.Name = "txtNumidOrigen1";
			this.txtNumidOrigen1.PromptChar = ' ';
			this.txtNumidOrigen1.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen1.TabIndex = 326;
			this.txtNumidOrigen1.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(719, 440);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 325;
			this.txtNumContadoCredito.Visible = false;
			this.txtNumContadoCredito.ValueChanged += new System.EventHandler(this.txtNumContadoCredito_ValueChanged);
			// 
			// txtNumidOrigen
			// 
			this.txtNumidOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumidOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumidOrigen.Enabled = false;
			this.txtNumidOrigen.Location = new System.Drawing.Point(559, 440);
			this.txtNumidOrigen.Name = "txtNumidOrigen";
			this.txtNumidOrigen.PromptChar = ' ';
			this.txtNumidOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtNumidOrigen.TabIndex = 324;
			this.txtNumidOrigen.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(543, 440);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 323;
			this.txtNumIdCompra.Visible = false;
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(671, 440);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 322;
			this.txtNumIdCliente.Visible = false;
			// 
			// txtNumIdAsiento
			// 
			this.txtNumIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdAsiento.Enabled = false;
			this.txtNumIdAsiento.Location = new System.Drawing.Point(639, 440);
			this.txtNumIdAsiento.Name = "txtNumIdAsiento";
			this.txtNumIdAsiento.PromptChar = ' ';
			this.txtNumIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdAsiento.TabIndex = 321;
			this.txtNumIdAsiento.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(687, 440);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoRuc.TabIndex = 335;
			this.txtIdTipoRuc.Visible = false;
			// 
			// btnRetencion
			// 
			this.btnRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnRetencion.CausesValidation = false;
			this.btnRetencion.Enabled = false;
			this.btnRetencion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRetencion.ImageIndex = 64;
			this.btnRetencion.ImageList = this.imageList1;
			this.btnRetencion.Location = new System.Drawing.Point(485, 488);
			this.btnRetencion.Name = "btnRetencion";
			this.btnRetencion.Size = new System.Drawing.Size(80, 23);
			this.btnRetencion.TabIndex = 25;
			this.btnRetencion.Text = "&Retención";
			this.btnRetencion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnRetencion, "Ver Retencion de la Factura Actual");
			this.btnRetencion.Click += new System.EventHandler(this.btnRetencion_Click);
			// 
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto1.DataSource = this.ultraDataSource2;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 200;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmbProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(376, 192);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(256, 112);
			this.cmbProyecto1.TabIndex = 337;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbTipoFormaPago
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFormaPago.Appearance = appearance39;
			this.cmbTipoFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbTipoFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipoFormaPago.DisplayMember = "FormaDePago";
			this.cmbTipoFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFormaPago.Enabled = false;
			this.cmbTipoFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFormaPago.Location = new System.Drawing.Point(96, 96);
			this.cmbTipoFormaPago.Name = "cmbTipoFormaPago";
			this.cmbTipoFormaPago.Size = new System.Drawing.Size(200, 22);
			this.cmbTipoFormaPago.TabIndex = 14;
			this.cmbTipoFormaPago.ValueMember = "idTipoFormaPago";
			this.cmbTipoFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoFormaPago_KeyDown);
			this.cmbTipoFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFormaPago_InitializeLayout);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 99);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(70, 16);
			this.label7.TabIndex = 339;
			this.label7.Text = "Tipo de pago";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbImportacion
			// 
			this.cmbImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbImportacion.DataSource = this.ultraDataSource3;
			this.cmbImportacion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand8.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Width = 278;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43});
			this.cmbImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbImportacion.DisplayMember = "Nombre";
			this.cmbImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbImportacion.Location = new System.Drawing.Point(552, 208);
			this.cmbImportacion.Name = "cmbImportacion";
			this.cmbImportacion.Size = new System.Drawing.Size(280, 112);
			this.cmbImportacion.TabIndex = 340;
			this.cmbImportacion.ValueMember = "idImportacionTipo";
			this.cmbImportacion.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(703, 440);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdGrupoCliente.TabIndex = 341;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// dtCaducidad
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtCaducidad.Appearance = appearance40;
			this.dtCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtCaducidad.DateButtons.Add(dateButton2);
			this.dtCaducidad.Enabled = false;
			this.dtCaducidad.Format = "dd/MM/yyyy";
			this.dtCaducidad.Location = new System.Drawing.Point(560, 416);
			this.dtCaducidad.Name = "dtCaducidad";
			this.dtCaducidad.NonAutoSizeHeight = 23;
			this.dtCaducidad.Size = new System.Drawing.Size(112, 21);
			this.dtCaducidad.SpinButtonsVisible = true;
			this.dtCaducidad.TabIndex = 9;
			this.dtCaducidad.Value = ((object)(resources.GetObject("dtCaducidad.Value")));
			this.dtCaducidad.Visible = false;
			this.dtCaducidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtCaducidad_KeyDown);
			this.dtCaducidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtCaducidad_KeyPress);
			this.dtCaducidad.Validating += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_Validating);
			this.dtCaducidad.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_BeforeDropDown);
			// 
			// txtIdTipoDePago
			// 
			this.txtIdTipoDePago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoDePago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoDePago.Enabled = false;
			this.txtIdTipoDePago.Location = new System.Drawing.Point(735, 440);
			this.txtIdTipoDePago.Name = "txtIdTipoDePago";
			this.txtIdTipoDePago.PromptChar = ' ';
			this.txtIdTipoDePago.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoDePago.TabIndex = 343;
			this.txtIdTipoDePago.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(495, 392);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 344;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(640, 68);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(92, 16);
			this.label29.TabIndex = 346;
			this.label29.Text = "Crédito Tributario";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCreditoTributario
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCreditoTributario.Appearance = appearance41;
			this.cmbCreditoTributario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCreditoTributario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCreditoTributario.DataSource = this.ultraDataSource5;
			ultraGridBand9.AddButtonCaption = "CompraCredTribut";
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 55;
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Width = 80;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Width = 312;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			this.cmbCreditoTributario.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbCreditoTributario.DisplayMember = "CredTributario";
			this.cmbCreditoTributario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCreditoTributario.Enabled = false;
			this.cmbCreditoTributario.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCreditoTributario.Location = new System.Drawing.Point(744, 65);
			this.cmbCreditoTributario.Name = "cmbCreditoTributario";
			this.cmbCreditoTributario.Size = new System.Drawing.Size(280, 22);
			this.cmbCreditoTributario.TabIndex = 16;
			this.cmbCreditoTributario.ValueMember = "idCredTributario";
			this.cmbCreditoTributario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCreditoTributario_KeyDown);
			this.cmbCreditoTributario.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCreditoTributario_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(640, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(71, 16);
			this.label9.TabIndex = 348;
			this.label9.Text = "Comprobante";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbComprobante
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComprobante.Appearance = appearance42;
			this.cmbComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DataSource = this.ultraDataSource6;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 264;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn47,
																											ultraGridColumn48});
			this.cmbComprobante.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbComprobante.DisplayMember = "Comprobante";
			this.cmbComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobante.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComprobante.Enabled = false;
			this.cmbComprobante.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobante.Location = new System.Drawing.Point(720, 5);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(304, 22);
			this.cmbComprobante.TabIndex = 347;
			this.cmbComprobante.ValueMember = "idComprobante";
			// 
			// ultraDataSource6
			// 
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn38,
																																 ultraDataColumn39});
			// 
			// lblOrigen
			// 
			this.lblOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblOrigen.AutoSize = true;
			this.lblOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblOrigen.ForeColor = System.Drawing.Color.Firebrick;
			this.lblOrigen.Location = new System.Drawing.Point(472, 432);
			this.lblOrigen.Name = "lblOrigen";
			this.lblOrigen.Size = new System.Drawing.Size(0, 26);
			this.lblOrigen.TabIndex = 349;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource7;
			appearance43.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance43;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 2;
			ultraGridColumn51.Width = 8;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn49,
																											ultraGridColumn50,
																											ultraGridColumn51});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance44;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance46.FontData.BoldAsString = "True";
			appearance46.FontData.Name = "Arial";
			appearance46.FontData.SizeInPoints = 10F;
			appearance46.ForeColor = System.Drawing.Color.White;
			appearance46.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance46;
			appearance47.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance47.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(832, 184);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(24, 80);
			this.ultraGrid2.TabIndex = 350;
			this.ultraGrid2.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42});
			// 
			// btnProforma
			// 
			this.btnProforma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnProforma.CausesValidation = false;
			this.btnProforma.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProforma.ImageIndex = 51;
			this.btnProforma.ImageList = this.imageList1;
			this.btnProforma.Location = new System.Drawing.Point(764, 488);
			this.btnProforma.Name = "btnProforma";
			this.btnProforma.Size = new System.Drawing.Size(76, 23);
			this.btnProforma.TabIndex = 356;
			this.btnProforma.Text = "&Proforma";
			this.btnProforma.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnProforma, "Busca Ordenes de Compra Procesadas");
			this.btnProforma.Click += new System.EventHandler(this.btnProforma_Click);
			// 
			// btnSubir
			// 
			this.btnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSubir.CausesValidation = false;
			this.btnSubir.Image = ((System.Drawing.Image)(resources.GetObject("btnSubir.Image")));
			this.btnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubir.Location = new System.Drawing.Point(472, 440);
			this.btnSubir.Name = "btnSubir";
			this.btnSubir.Size = new System.Drawing.Size(56, 23);
			this.btnSubir.TabIndex = 359;
			this.btnSubir.Text = "&Subir";
			this.btnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSubir, "Busca Ordenes de Compra Procesadas");
			this.btnSubir.Visible = false;
			this.btnSubir.Click += new System.EventHandler(this.btnSubir_Click);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn43.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45});
			// 
			// cmbProyectoZOOM
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyectoZOOM.Appearance = appearance50;
			this.cmbProyectoZOOM.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyectoZOOM.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyectoZOOM.DataSource = this.ultraDataSource8;
			ultraGridBand12.AddButtonCaption = "CompraCredTribut";
			ultraGridColumn52.Header.VisiblePosition = 0;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 1;
			ultraGridColumn54.Header.VisiblePosition = 2;
			ultraGridColumn54.Width = 300;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn52,
																											ultraGridColumn53,
																											ultraGridColumn54});
			this.cmbProyectoZOOM.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbProyectoZOOM.DisplayMember = "Nombre";
			this.cmbProyectoZOOM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyectoZOOM.Enabled = false;
			this.cmbProyectoZOOM.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyectoZOOM.Location = new System.Drawing.Point(392, 65);
			this.cmbProyectoZOOM.Name = "cmbProyectoZOOM";
			this.cmbProyectoZOOM.Size = new System.Drawing.Size(232, 22);
			this.cmbProyectoZOOM.TabIndex = 352;
			this.cmbProyectoZOOM.ValueMember = "idProyectoZoom";
			this.cmbProyectoZOOM.Visible = false;
			// 
			// cmbTiposDeGastos
			// 
			this.cmbTiposDeGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposDeGastos.DataSource = this.ultraDataSource9;
			this.cmbTiposDeGastos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand13.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 121;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Width = 278;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn55,
																											ultraGridColumn56});
			this.cmbTiposDeGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbTiposDeGastos.DisplayMember = "Descripción";
			this.cmbTiposDeGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposDeGastos.Location = new System.Drawing.Point(72, 184);
			this.cmbTiposDeGastos.Name = "cmbTiposDeGastos";
			this.cmbTiposDeGastos.Size = new System.Drawing.Size(280, 112);
			this.cmbTiposDeGastos.TabIndex = 353;
			this.cmbTiposDeGastos.ValueMember = "idTipoGasto";
			this.cmbTiposDeGastos.Visible = false;
			this.cmbTiposDeGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTiposDeGastos_InitializeLayout);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn46.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn46,
																																 ultraDataColumn47});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Enabled = false;
			this.label10.Location = new System.Drawing.Point(536, 128);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(78, 16);
			this.label10.TabIndex = 355;
			this.label10.Text = "Días Crédito 4:";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasCredito4
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito4.Appearance = appearance51;
			this.txtDiasCredito4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito4.Enabled = false;
			this.txtDiasCredito4.Location = new System.Drawing.Point(624, 125);
			this.txtDiasCredito4.MinValue = 0;
			this.txtDiasCredito4.Name = "txtDiasCredito4";
			this.txtDiasCredito4.PromptChar = ' ';
			this.txtDiasCredito4.Size = new System.Drawing.Size(64, 22);
			this.txtDiasCredito4.TabIndex = 354;
			this.txtDiasCredito4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito4_KeyDown);
			this.txtDiasCredito4.Click += new System.EventHandler(this.txtDiasCredito4_Click);
			this.txtDiasCredito4.Enter += new System.EventHandler(this.txtDiasCredito4_Enter);
			// 
			// lblFechaPEriodo
			// 
			this.lblFechaPEriodo.AutoSize = true;
			this.lblFechaPEriodo.Location = new System.Drawing.Point(472, 360);
			this.lblFechaPEriodo.Name = "lblFechaPEriodo";
			this.lblFechaPEriodo.Size = new System.Drawing.Size(75, 16);
			this.lblFechaPEriodo.TabIndex = 358;
			this.lblFechaPEriodo.Text = "Fecha Periodo";
			this.lblFechaPEriodo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblFechaPEriodo.Visible = false;
			// 
			// dtFechaPeriodo
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaPeriodo.Appearance = appearance52;
			this.dtFechaPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaPeriodo.DateButtons.Add(dateButton3);
			this.dtFechaPeriodo.Enabled = false;
			this.dtFechaPeriodo.Format = "dd/MM/yyyy";
			this.dtFechaPeriodo.Location = new System.Drawing.Point(560, 360);
			this.dtFechaPeriodo.Name = "dtFechaPeriodo";
			this.dtFechaPeriodo.NonAutoSizeHeight = 23;
			this.dtFechaPeriodo.Size = new System.Drawing.Size(112, 21);
			this.dtFechaPeriodo.SpinButtonsVisible = true;
			this.dtFechaPeriodo.TabIndex = 357;
			this.dtFechaPeriodo.Value = ((object)(resources.GetObject("dtFechaPeriodo.Value")));
			this.dtFechaPeriodo.Visible = false;
			this.dtFechaPeriodo.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaPeriodo_BeforeDropDown);
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(568, 392);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(72, 16);
			this.linkLabel1.TabIndex = 360;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "Documento";
			this.linkLabel1.Visible = false;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Enabled = false;
			this.label12.Location = new System.Drawing.Point(696, 128);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(78, 16);
			this.label12.TabIndex = 362;
			this.label12.Text = "Días Crédito 5:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasCredito5
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito5.Appearance = appearance53;
			this.txtDiasCredito5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito5.Enabled = false;
			this.txtDiasCredito5.Location = new System.Drawing.Point(784, 125);
			this.txtDiasCredito5.MinValue = 0;
			this.txtDiasCredito5.Name = "txtDiasCredito5";
			this.txtDiasCredito5.PromptChar = ' ';
			this.txtDiasCredito5.Size = new System.Drawing.Size(64, 22);
			this.txtDiasCredito5.TabIndex = 361;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(816, 38);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(91, 16);
			this.label13.TabIndex = 364;
			this.label13.Text = "Fecha de Emisión";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Click += new System.EventHandler(this.label13_Click);
			// 
			// dtFechadeEmicion
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechadeEmicion.Appearance = appearance54;
			this.dtFechadeEmicion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechadeEmicion.DateButtons.Add(dateButton4);
			this.dtFechadeEmicion.Enabled = false;
			this.dtFechadeEmicion.Format = "dd/MM/yyyy";
			this.dtFechadeEmicion.Location = new System.Drawing.Point(920, 36);
			this.dtFechadeEmicion.Name = "dtFechadeEmicion";
			this.dtFechadeEmicion.NonAutoSizeHeight = 23;
			this.dtFechadeEmicion.Size = new System.Drawing.Size(120, 21);
			this.dtFechadeEmicion.SpinButtonsVisible = true;
			this.dtFechadeEmicion.TabIndex = 363;
			this.dtFechadeEmicion.Value = ((object)(resources.GetObject("dtFechadeEmicion.Value")));
			this.dtFechadeEmicion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechadeEmicion_KeyDown);
			this.dtFechadeEmicion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechadeEmicion_KeyPress);
			this.dtFechadeEmicion.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechadeEmicion_Validating);
			// 
			// btnDistribucion
			// 
			this.btnDistribucion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDistribucion.CausesValidation = false;
			this.btnDistribucion.Enabled = false;
			this.btnDistribucion.Image = ((System.Drawing.Image)(resources.GetObject("btnDistribucion.Image")));
			this.btnDistribucion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDistribucion.Location = new System.Drawing.Point(943, 96);
			this.btnDistribucion.Name = "btnDistribucion";
			this.btnDistribucion.Size = new System.Drawing.Size(104, 24);
			this.btnDistribucion.TabIndex = 365;
			this.btnDistribucion.Text = "Distribucion";
			this.btnDistribucion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDistribucion.Click += new System.EventHandler(this.btnDistribucion_Click);
			// 
			// mnuDistribucion
			// 
			this.mnuDistribucion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAgencias,
																																										this.mnuAdministrativos,
																																										this.mnuGeneral,
																																										this.mnuAgenciasI,
																																										this.mnuAdministrativosI,
																																										this.mnuGeneralI});
			this.mnuDistribucion.Popup += new System.EventHandler(this.mnuDistribucion_Popup);
			// 
			// mnuAgencias
			// 
			this.mnuAgencias.Index = 0;
			this.mnuAgencias.Text = "Agencias";
			this.mnuAgencias.Click += new System.EventHandler(this.mnuAgencias_Click);
			// 
			// mnuAdministrativos
			// 
			this.mnuAdministrativos.Index = 1;
			this.mnuAdministrativos.Text = "Administrativos";
			this.mnuAdministrativos.Click += new System.EventHandler(this.mnuAdministrativos_Click);
			// 
			// mnuGeneral
			// 
			this.mnuGeneral.Index = 2;
			this.mnuGeneral.Text = "General";
			this.mnuGeneral.Click += new System.EventHandler(this.mnuGeneral_Click);
			// 
			// mnuAgenciasI
			// 
			this.mnuAgenciasI.Index = 3;
			this.mnuAgenciasI.Text = "Agencias -Igualdad";
			this.mnuAgenciasI.Click += new System.EventHandler(this.mnuAgenciasI_Click);
			// 
			// mnuAdministrativosI
			// 
			this.mnuAdministrativosI.Index = 4;
			this.mnuAdministrativosI.Text = "Administrativos - Igualdad";
			this.mnuAdministrativosI.Click += new System.EventHandler(this.mnuAdministrativosI_Click);
			// 
			// mnuGeneralI
			// 
			this.mnuGeneralI.Index = 5;
			this.mnuGeneralI.Text = "General - Igualdad";
			this.mnuGeneralI.Click += new System.EventHandler(this.mnuGeneralI_Click);
			// 
			// btnTemporal
			// 
			this.btnTemporal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnTemporal.CausesValidation = false;
			this.btnTemporal.Enabled = false;
			this.btnTemporal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTemporal.ImageIndex = 66;
			this.btnTemporal.ImageList = this.imageList1;
			this.btnTemporal.Location = new System.Drawing.Point(686, 488);
			this.btnTemporal.Name = "btnTemporal";
			this.btnTemporal.TabIndex = 366;
			this.btnTemporal.Text = "Temporal";
			this.btnTemporal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTemporal.Click += new System.EventHandler(this.btnTemporal_Click);
			// 
			// btnDias
			// 
			this.btnDias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDias.Enabled = false;
			this.btnDias.Image = ((System.Drawing.Image)(resources.GetObject("btnDias.Image")));
			this.btnDias.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDias.Location = new System.Drawing.Point(807, 96);
			this.btnDias.Name = "btnDias";
			this.btnDias.Size = new System.Drawing.Size(112, 24);
			this.btnDias.TabIndex = 1010;
			this.btnDias.Text = "Días de Crédito";
			this.btnDias.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDias.Visible = false;
			this.btnDias.Click += new System.EventHandler(this.btnDias_Click);
			// 
			// frmCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1055, 518);
			this.Controls.Add(this.btnDias);
			this.Controls.Add(this.btnTemporal);
			this.Controls.Add(this.btnDistribucion);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.lblFechaPEriodo);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblImportacion);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label69);
			this.Controls.Add(this.label70);
			this.Controls.Add(this.label71);
			this.Controls.Add(this.dtFechadeEmicion);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.btnSubir);
			this.Controls.Add(this.dtFechaPeriodo);
			this.Controls.Add(this.btnProforma);
			this.Controls.Add(this.cmbTiposDeGastos);
			this.Controls.Add(this.cmbProyectoZOOM);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.lblOrigen);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.cmbCreditoTributario);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtIdTipoDePago);
			this.Controls.Add(this.dtCaducidad);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.cmbImportacion);
			this.Controls.Add(this.cmbTipoFormaPago);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.btnRetencion);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtNumIdproyectoZoom);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumidOrigen1);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumidOrigen);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtNumIdAsiento);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnBuscarOrdenDeCompra);
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
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.cmbImportacionNombre);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtDiasCredito5);
			this.Controls.Add(this.txtDiasCredito4);
			this.Controls.Add(this.txtDiasCredito3);
			this.Controls.Add(this.txtDiasCredito2);
			this.Controls.Add(this.txtDiasCredito1);
			this.KeyPreview = true;
			this.Name = "frmCompra";
			this.Text = "Compras";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCompra_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmCompra_Closing);
			this.Load += new System.EventHandler(this.frmCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacionNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdproyectoZoom)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumidOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoDePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyectoZOOM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechadeEmicion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

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

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
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

		private void UnloadMe()
		{
			this.Close();
		}
		
		DateTime dtFechaAFin = DateTime.Today;
		DateTime dtFechaAIni = DateTime.Today;
		DateTime dtFechaRecepcion = DateTime.Now;
		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";

		private void Consultar(int idCompra)
		{
			try
			{
				#region Maestro
					
				/*c.idCompra, c.idOrigen, c.idSucursal, c.idTipoFactura, c.idAsiento, c.idComprobante, c.idCredTributario, c.Fiscal, c.idMotivo, (0-8)
					c.Bodega, c.BodegaDestino, c.idProyecto, c.idProyectoZoom, (9-12)
					c.ContadoCredito, c.idFormaPago, c.idTipoDePago, c.idTipoFormaPago, (13-16)
					c.Fecha, c.Numero, c.Comprobante, c.SerieFactura, c.AutFactura, c.FechaCaducidad, (17-22)
					c.idCliente, cl.idTipoRuc, cl.idGrupoCliente, cl.Ruc, cl.Nombre, cl.DiasCredito, cl.DiasCredito1, cl.DiasCredito2, (23-30)
					c.CrediPoint, c.FechaPrimerPago, c.TotalCuotas, c.NumCuotas, c.CuotaEntrada, c.ValorCuotas, (31-36)
					c.idImportacionTipo, c.idVendedor, (37-38)
					c.SubtotalExcento, c.Descuento0, c.SubtotalIva, c.Descuento, c.Iva, c.Total, c.Usuario, (39-45) 
					c.idEntidadFinanciera,  c.CuotaAsignada, c.CuotaDisponible, c.CuotaUsada, c.Interes, (46-50)
					c.Observaciones, c.Notas, (51-52)
					c.Entregado, c.FechaEntrega, c.HoraEntrega, c.idDireccion, (53-56)
					c.Estado, c.Impreso, c.Borrar, c.Consignacion (57-60) */

				string sSQLBusca = string.Format("Exec ConsultaCompraIndividual {0}", idCompra);
				SqlDataReader drBusca = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusca, true);
				drBusca.Read();
				if (drBusca.HasRows)
				{					
					this.txtNumIdCompra.Value = drBusca.GetInt32(0);
					if (drBusca.GetValue(1) != System.DBNull.Value) this.txtNumidOrigen.Value = drBusca.GetInt32(1);
					this.txtIdSucursal.Value  = drBusca.GetInt32(2);
					this.txtNumEstado.Value = drBusca.GetInt32(57);
					if (drBusca.GetValue(4) != System.DBNull.Value) this.txtNumIdAsiento.Value = drBusca.GetInt32(4);
					if (drBusca.GetValue(5) != System.DBNull.Value) this.cmbComprobante.Value = drBusca.GetInt32(5);
					this.txtNumContadoCredito.Value = drBusca.GetInt32(13);
					this.cmbFormaPago.Value = drBusca.GetInt32(14);
					this.cmbBodega.Value = drBusca.GetInt32(9);
					this.cmbProyecto.Value = drBusca.GetInt32(11);
					this.cmbProyectoZOOM.Value = drBusca.GetInt32(12);
					this.txtNumero.Text = drBusca.GetString(18);
					if (drBusca.GetValue(19)  != System.DBNull.Value) this.txtComprobante.Text = drBusca.GetString(19);
					this.txtIdTipoRuc.Value = drBusca.GetInt32(24);
					this.txtNumIdCliente.Value = drBusca.GetInt32(23);
					if (drBusca.GetValue(25) != System.DBNull.Value) this.txtIdGrupoCliente.Value = drBusca.GetInt32(25);
					if (drBusca.GetValue(37) != System.DBNull.Value) this.cmbImportacionNombre.Value = drBusca.GetInt32(37);
					if (drBusca.GetValue(16) != System.DBNull.Value) this.cmbTipoFormaPago.Value = drBusca.GetInt32(16);
					if (drBusca.GetValue(6) != System.DBNull.Value) this.cmbCreditoTributario.Value = drBusca.GetInt32(6);
					if (drBusca.GetValue(20) != System.DBNull.Value) this.txtSerieFact.Text = drBusca.GetString(20).ToString();
					if (drBusca.GetValue(21) != System.DBNull.Value) this.txtAutFactura.Text = drBusca.GetString(21).ToString();
					if (drBusca.GetValue(22) != System.DBNull.Value) this.dtCaducidad.Value = drBusca.GetDateTime(22);
					if (drBusca.GetValue(15) != System.DBNull.Value) this.txtIdTipoDePago.Value = drBusca.GetInt32(15);
					this.txtRuc.Text = drBusca.GetString(26).ToString();
					this.txtNombre.Text = drBusca.GetString(27).ToString();
					this.txtDiasCredito1.Value = drBusca.GetInt32(28);
					this.txtDiasCredito2.Value = drBusca.GetInt32(29);
					this.txtDiasCredito3.Value = drBusca.GetInt32(30);
					this.txtDiasCredito4.Value = drBusca.GetInt32(63);
					this.txtDiasCredito5.Value = drBusca.GetInt32(67);
					this.dtFecha.Value = drBusca.GetDateTime(17);
					if (drBusca.GetValue(52) != System.DBNull.Value) this.txtNotas.Text = drBusca.GetString(52).ToString();
					this.txtIva0.Value = drBusca.GetDecimal(39);
					this.txtDescIva0Total.Value = drBusca.GetDecimal(40);
					this.txtIva.Value = drBusca.GetDecimal(41);
					this.txtDescIvaTotal.Value = drBusca.GetDecimal(42);
					this.txtIvaTotal.Value = drBusca.GetDecimal(43);
					this.txtTotal.Value = drBusca.GetDecimal(44);
					bImpreso = drBusca.GetBoolean(58);
					bAnulado = drBusca.GetBoolean(59);
					this.chkConsignacion.Checked = drBusca.GetBoolean(60);
					iOrigen = drBusca.GetInt32(61);
					if (drBusca.GetValue(64) != System.DBNull.Value) this.dtFechaPeriodo.Value = drBusca.GetDateTime(64);
					this.dtFechadeEmicion.Value = drBusca.GetDateTime(68);
				
					
					this.lblUsuario.Text = drBusca.GetString(45).ToString();

					if (bAnulado) this.lblEstado.Text = "ANULADO";
					else
					{
						if ((int)this.txtNumEstado.Value == 7) this.lblEstado.Text = "POR LIQUIDAR";
						if ((int)this.txtNumEstado.Value == 6) this.lblEstado.Text = "TEMPORAL";
						if ((int)this.txtNumEstado.Value == 8) this.lblEstado.Text = "LIQUIDADO";
						if ((int)this.txtNumEstado.Value == 12) this.lblEstado.Text = "FACTURADO SIN MERCADERIA";
						if ((int)this.txtNumEstado.Value == 14) this.lblEstado.Text = "LIQUIDADO";									  
					}

					//					if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33 )
					//					{
					if (drBusca.GetValue(65) != System.DBNull.Value) sCarpeta = drBusca.GetString(65).ToString();
					if (drBusca.GetValue(66) != System.DBNull.Value) sUrlNube = drBusca.GetString(66).ToString();
					if (sUrlNube.Length > 10 )
					{
						this.linkLabel1.Visible = true;
						//this.btnSubir.Visible = true;
					}
					//	}
					
				}
				drBusca.Close();
				#endregion Maestro 					

				if (iOrigen == 1) this.lblOrigen.Text = "CAJA CHICA";
				if (iOrigen == 2) this.lblOrigen.Text = "TARJETA DE CRÉDITO";
 
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCompra {0}", (int)this.txtNumIdCompra.Value));											
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaODCCompra {0}", (int)this.txtNumIdCompra.Value));

				if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					#region Editar
					if (miAcceso.Editar)
					{
						if ((int)this.txtNumEstado.Value == 8)
						{
							if (miAcceso.BEditarLiquidado) this.btnEditar.Enabled = true;	
						}
						else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 0 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
						{
							if (miAcceso.BEditarPagosRealizados) this.btnEditar.Enabled = true;	
						}
						else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos Where Estado = 1 And idTipo = 4 And idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
						{
							if (miAcceso.BEditarPagosAutorizados)	this.btnEditar.Enabled = true;
						}
						else this.btnEditar.Enabled = true;
					}
					
			

					#endregion Editar	

					#region Ediat del dia
					DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
					DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
					if (miAcceso.EditarSoloDia)
					{

						if (dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString() )
						{
							if ((int)this.txtNumEstado.Value == 8)
							{
								if (miAcceso.BEditarLiquidado) this.btnEditar.Enabled = true;	
							}
							else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 0 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
							{
								if (miAcceso.BEditarPagosRealizados) this.btnEditar.Enabled = true;	
							}
							else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos Where Estado = 1 And idTipo = 4 And idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
							{
								if (miAcceso.BEditarPagosAutorizados)	this.btnEditar.Enabled = true;
							}
							else this.btnEditar.Enabled = true;
						}
					}
					#endregion Ediat del dia 
					
					#region Anular
					if (!bAnulado && miAcceso.Anular)
					{
						if ((int)this.txtNumEstado.Value == 8)
						{
							if (miAcceso.BAnularLiquidado) this.btnAnular.Enabled = true;	
						}
						else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 1 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
						{
							if (miAcceso.BAnularPagosRealizados) this.btnAnular.Enabled = true;	
						}
						else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos Where Estado = 1 And idTipo = 4 And idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
						{
							if (miAcceso.BAnularPagosAutorizados) this.btnAnular.Enabled = true;
						}
						else this.btnAnular.Enabled = true;
					}
					#endregion Anular

					if (bAnulado && miAcceso.Eliminar) this.btnAnular.Enabled = true;
				}
				else
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("El Período de esta Factura '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);								
				}

				if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;			
				this.btnCancelar.Enabled = true;
				this.btnDias.Visible = true;
				this.btnDias.Enabled = true;

				if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
				if (miAcceso.Retencion) this.btnRetencion.Enabled = true;
				if (bAnulado) this.btnAnular.Text = "Eliminar";

				#region Liquidacion Compra Validacion
				int RetiLiqCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From LiquidacionCompra Where idcompra = {0}", idCompra));
				if (RetiLiqCompra == 1)
				{
					bLiquidacion = true;
				}
				else
				{
					bLiquidacion = false;
				}
				#endregion Liquidacion Compra Validacion
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		int idAgenteRetencion = 0;

		private void frmCompra_Load(object sender, System.EventArgs e)
		{													
			miAcceso = new Acceso(cdsSeteoF, "0603");
			
			if (!Funcion.Permiso("103", cdsSeteo))
			{				
				MessageBox.Show("No puede ingresar a Compras", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;

			ModificaGrilla(false);

			Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'COMPRA'", true);
			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");
			bAsientoAutomatico = Funcion.bEjecutaSQL(cdsSeteoF, "Exec AsientoAutomaticoVerifica 4");
			IdSigno = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 4");
			idAgenteRetencion = Funcion.iEscalarSQL(cdsSeteoF, "Select idAgenteRetencion From Seteo");
			
			#region Lectura de Archivo - Propiedades - Configuracion
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];				
			#endregion Lectura de Archivo - Propiedades - Configuracion

			dtFechaAFin = (DateTime)drSeteoC["FechaFin"];
			dtFechaAIni = (DateTime)drSeteoC["FechaIni"];
			bGastosPorProyecto = (bool)drSeteoF["GastosPorProyecto"];

			if (bGastosPorProyecto)
			{
				this.lblImportacion.Text = "Proyecto";
				this.cmbProyectoZOOM.Visible = true;
				this.cmbProyectoZOOM.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyectoZoom, Numero, Nombre From ProyectoZoom Order By Nombre");
				this.cmbTiposDeGastos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoGasto, Descripción From TiposDeGasto Order By Descripción");
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idTipoGasto"].Hidden = false;
				this.cmbImportacionNombre.Visible = false;
			}

			#region Listas
			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 4, DateTime.Today, false, cdsSeteoF);
			//this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BodegaUsuario 4");
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ProyectoCarga 4");
			this.cmbProyecto1.DataSource = cmbProyecto.DataSource;
			this.cmbComprobante.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idComprobante, Comprobante From CompraComprobante Where idComprobante In (1, 2, 3, 12) Order By Codigo");
			
			if (miAcceso.EditarDetalleFormaPago) 
			{
				#region Valida Empresa
				//				if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 3)
				//				{
				//					this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCargaCredito");
				//				}
				//				else
				//				{
				//					this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga 4");
				//				}
				#endregion Valida Empresa

				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCargaCredito");
			}
			else
			{
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga 4");
			}

			this.cmbTipoFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFormaPago, FormaDePago From TipoFormaDePago Where Activo = 1");
			this.cmbImportacionNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idImportacionNombre, Nombre From ImportacionNombre Order By idImportacionNombre");
			this.cmbImportacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idImportacionTipo, Nombre From ImportacionTipo Where LEN(Nombre) > 0 Order By idImportacionTipo");
			this.cmbCreditoTributario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCredTributario, Codigo, CredTributario From CompraCredTribut Order By Codigo");
			#endregion Listas

			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 4, 3", true);
			this.txtIdTipoFactura.Value = 4;			
			this.txtIdSucursal.Value = 1;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			//this.dtFechadeEmicion.CalendarInfo.MinDate = DateTime.Today.AddDays(-31);
			this.dtFechadeEmicion.CalendarInfo.MaxDate = DateTime.Today;
			//this.dtFechaPeriodo.CalendarInfo.MaxDate = DateTime.Today;
			//this.dtFecha.CalendarInfo.MinDate=DateTime.Parse("28/02/2022");

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
			if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
			
			int iOrigen = 0;

			#region Abre Registro
			if (BBuscar) this.Consultar(IdCompra);
			if (IdCompra>0) 
			{
				this.btnDias.Visible=true;
				this.btnDias.Enabled=true;

			}
			if (bConsignacion) 
			{
				this.btnNuevo_Click(sender, e);

				this.txtRuc.Text = sCSRUC;
				
				this.txtRuc_Validated(sender, e);

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec CSListaProductosFacturaProveedor {0}, {1}", (int)this.txtNumIdCliente.Value, iCSidFormaPago));				

				this.txtRuc.Enabled = false;
			}
			#endregion Abre Registro

			if (BEditar) this.btnEditar_Click(sender, e);

			if (MenuLatinium.stDirFacturacion == "DENTALIS")
			{
				this.dtFechaPeriodo.Visible = true;
				this.lblFechaPEriodo.Visible = true;
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

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

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			if (bEdicion)
			{				
				if ((int)this.txtNumidOrigen.Value > 0)
				{
					string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueoOrigen, true);					
				}

				if ((int)this.txtNumidOrigen1.Value > 0)
				{
					string sSQLBloqueoOrigen1 = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueoOrigen1, true);
				}				
			}

			bEdicion = false;
			bNuevo = false;


			this.btnNuevo.Focus();

			this.txtNumIdCompra.Value = 0;
			this.txtNumidOrigen.Value = 0;
			this.txtNumidOrigen1.Value = 0;
			this.txtNumIdCliente.Value = 0;
			this.txtNumContadoCredito.Value = 0;
			this.txtNumIdAsiento.Value = 0;
			this.txtNumIdproyectoZoom.Value = 0;
			this.txtNumEstado.Value = 0;
			this.txtIdTipoRuc.Value = 0;
			this.txtIdGrupoCliente.Value = 0;
			this.txtIdTipoDePago.Value = 0;	
			idSujetoRetencion = 0;

			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbProyecto.Value = System.DBNull.Value;
			this.txtSerieFact.Text = "";
			this.txtNumero.Text = "";
			this.txtAutFactura.Text = "";
			this.dtCaducidad.Value = System.DBNull.Value;
			
			this.txtComprobante.Text = "";

			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtDiasCredito1.Value = 0;
			this.txtDiasCredito2.Value = 0;
			this.txtDiasCredito3.Value = 0;
			this.txtDiasCredito4.Value = 0;
			this.txtDiasCredito5.Value = 0;
			this.cmbComprobante.Value = System.DBNull.Value;
			this.cmbTipoFormaPago.Value = System.DBNull.Value;
			this.chkConsignacion.Checked = false;
			this.cmbImportacionNombre.Value = System.DBNull.Value;
			this.cmbProyectoZOOM.Value = System.DBNull.Value;
			this.cmbCreditoTributario.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechadeEmicion.Value = System.DBNull.Value;
			this.dtFechaPeriodo.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtFechaPeriodo.Value = System.DBNull.Value;
			

			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.lblOrigen.Text = "";
			
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleCompra 0");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaODCCompra 0");

			this.txtNotas.Text = "";

			this.txtIva0.Value = 0;
			this.txtIva.Value = 0;
			this.txtDescIva0Total.Value = 0;
			this.txtDescIvaTotal.Value = 0;
			this.txtIvaTotal.Value = 0;
			this.txtTotal.Value = 0;
			
			this.txtRuc.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.txtSerieFact.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtAutFactura.Enabled = false;
			this.dtCaducidad.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtDiasCredito1.Enabled = false;
			this.txtDiasCredito2.Enabled = false;
			this.txtDiasCredito3.Enabled = false;
			this.txtDiasCredito4.Enabled = false;
			this.txtDiasCredito5.Enabled = false;
			this.cmbComprobante.Enabled = false;
			this.cmbTipoFormaPago.Enabled = false;
			this.cmbImportacionNombre.Enabled = false;
			this.cmbProyectoZOOM.Enabled = false;
			this.cmbCreditoTributario.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechadeEmicion.Enabled = false;
			this.dtFechaPeriodo.Enabled = false;
			this.btnDistribucion.Enabled = false;
			
			sUrlNube ="";
			sCarpeta ="";
			txtDirectorioarchivo ="";
			this.btnSubir.Visible = false;
			this.linkLabel1.Visible = false;
			
			ModificaGrilla(false);

			this.txtNotas.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
			this.btnProforma.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnTemporal.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnRetencion.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnDistribucion.Enabled = false;

			this.btnAnular.Text = "Anular";
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["idImportacion"].Hidden = true;

			#region Variables
			iOrigen = 0;
			bEliminaAlValidar = false;
			bImpreso = false;			
			bAnulado = false;
			idBloqueaTransacciones = 0;
			dPorcIva = 0.00m;
			dtFechaRecepcion = DateTime.Today;
			bProforma = false;
			dValorProforma = 0.00m;
			bLiquidacion = false;
			#endregion Variables

			if (bConsignacion) this.DialogResult = DialogResult.No;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.btnCancelar_Click(sender, e);

			this.txtRuc.Enabled = true;
			this.txtSerieFact.Enabled = true;
			this.txtNumero.Enabled = true;
			this.txtAutFactura.Enabled = true;
			this.dtCaducidad.Enabled = true;
			this.dtFechadeEmicion.Enabled = true;
			this.dtFechaPeriodo.Enabled = true;
			this.txtDiasCredito1.Enabled = false;
			this.txtDiasCredito2.Enabled = false;
			this.txtDiasCredito3.Enabled = false;
			this.txtDiasCredito4.Enabled = false;
			this.txtDiasCredito5.Enabled = false;
			this.cmbTipoFormaPago.Enabled = true;
			if (miAcceso.BtnDistribucion) this.btnDistribucion.Enabled = true;
			if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
			
			//if ((int)MenuLatinium.stIdDB !=1 && (int)MenuLatinium.stIdDB !=33 )
			//	this.btnSubir.Visible = true;
				
			//			if (MenuLatinium.stDirFacturacion != "POINT" && MenuLatinium.stDirFacturacion != "CREDISOLUTIONS" ) 
			//			{
			//				this.btnSubir.Visible = true;
			//			}
			
			if (bGastosPorProyecto) this.cmbProyectoZOOM.Enabled = true; else this.cmbImportacionNombre.Enabled = true;

			this.cmbCreditoTributario.Enabled = true;
			//if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
			DateTime dtFechaVal1 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT DATEADD(DAY,3,EOMONTH('{0}',-1))",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));
			DateTime dtFechaVal2 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT EOMONTH('{0}',-1)",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));	
			if (miAcceso.BFecha3Días) 
			{
				int iCan= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDiasCompas '{0}','{1}'", Convert.ToDateTime(dtFechaVal2).ToString("yyyyMMdd"),Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));
				if(iCan<4)
				{
					this.dtFecha.Enabled = true;
					this.dtFecha.CalendarInfo.MinDate = dtFechaVal2;
					this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(iCan);
				}
					
			}
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			
			ModificaGrilla(true);
			
			this.txtNotas.Enabled = true;
						
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnGuardar.Enabled = true;
			if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnBuscarOrdenDeCompra.Enabled = false;
			this.btnProforma.Enabled = false;

			this.cmbFormaPago.Value = 6;
			this.txtNumContadoCredito.Value = 2;
			this.txtNumEstado.Value = 7;
			this.cmbTipoFormaPago.Value = 20;

				
			if (miAcceso.EditarDetalleFormaPago) 
			{
				#region Valida Empresa
				//				if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 3)
				//				{
				//					this.cmbFormaPago.Enabled = true;
				//				}
				//				else
				//				{
				//          this.cmbFormaPago.Enabled = false;
				//				}
				#endregion Valida Empresa
				this.cmbFormaPago.Enabled = true;
			}
			else
			{
				this.cmbFormaPago.Enabled = false;
			}

	

			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			//this.dtFecha.CalendarInfo.MinDate=DateTime.Parse("28/02/2022");

			this.txtRuc.Focus();

			bNuevo = true;
			bEdicion = true;

			this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
		int idSujetoRetencion = 0;

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			#region Busca Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, (bool)this.chkConsignacion.Checked);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (drC.HasRows)
			{
				this.txtNumIdCliente.Value = (int) drC.GetValue(0);
				this.txtRuc.Text = (string) drC.GetValue(1);
				this.txtNombre.Text = (string) drC.GetValue(2);
				this.txtIdGrupoCliente.Value = (int) drC.GetValue(3);						
				this.txtNumContadoCredito.Value = (int) drC.GetValue(6);
				this.txtIdTipoRuc.Value = (int) drC.GetValue(7);
				idSujetoRetencion = (int) drC.GetValue(10);
			}
			drC.Close();	
			#region Valida Venta Verde por liquidar

			int idProveedorVV = 0;

			if (this.txtNumIdCliente.Value != null && this.txtNumIdCliente.Value != System.DBNull.Value)
			{
				idProveedorVV = Convert.ToInt32(this.txtNumIdCliente.Value);
			}

			if (idProveedorVV > 0)
			{
				int cantidadPorLiquidar = Funcion.iEscalarSQL(
					cdsSeteoF,
					string.Format("Exec Web_ProveedorVentaVerdePorLiquidar {0}", idProveedorVV),
					true
					);

				if (cantidadPorLiquidar > 0)
				{
					DialogResult respuesta = MessageBox.Show(
						"El proveedor Venta Verde tiene productos por liquidar.\n\n¿Deseas verlos?",
						"Venta Verde",
						MessageBoxButtons.YesNo,
						MessageBoxIcon.Question
						);

					if (respuesta == DialogResult.Yes)
					{
						frmLiquidarVentaVerde frm = new frmLiquidarVentaVerde(idProveedorVV);

						if (frm.ShowDialog() == DialogResult.OK)
						{
							if (frm.ProductosSeleccionados != null && frm.ProductosSeleccionados.Rows.Count > 0)
							{
								CargarVentaVerdeEnDetalle(frm.ProductosSeleccionados);
							}
						}
					}
				}
			}

			#endregion Valida Venta Verde por liquidar



			#endregion Busca Proveedor							
				
			bool bRise = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(InstPublica, 0) From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value), true);

			if ((int)this.txtIdTipoRuc.Value == 1) 
			{
				if (bRise) this.cmbComprobante.Value = 2;
				else this.cmbComprobante.Value = 1;				
			}
			else if ((int)this.txtIdTipoRuc.Value == 2) this.cmbComprobante.Value = 3;
			else if ((int)this.txtIdTipoRuc.Value == 3 && !bRise) 
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, "Select ImportacionAduana From SeteoF", true)) this.cmbComprobante.Value = 12;
				else this.cmbComprobante.Value = 3;
			}
			else if ((int)this.txtIdTipoRuc.Value == 3 && bRise) 
			{
				this.txtAutFactura.Enabled = true;
				this.txtSerieFact.Enabled = true;
				this.dtCaducidad.Enabled = true;
				this.dtFechadeEmicion.Enabled = true;
				this.dtFechaPeriodo.Enabled = true;
				this.cmbComprobante.Value = 3;
			}

			#region Dias de credito
			if ((int)this.txtNumContadoCredito.Value == 2)
			{
				string sSQLCredClte = string.Format("Select DiasCredito, DiasCredito1, DiasCredito2,  DiasCredito3, ISNULL(DiasCredito4,0) AS DiasCredito4  From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value);
				SqlDataReader drCredCliente = Funcion.rEscalarSQL(cdsSeteoF, sSQLCredClte, true);
				drCredCliente.Read();
				if (drCredCliente.HasRows)
				{
					this.txtDiasCredito1.Value = (int)drCredCliente.GetInt32(0);
					this.txtDiasCredito2.Value = (int)drCredCliente.GetInt32(1);
					this.txtDiasCredito3.Value = (int)drCredCliente.GetInt32(2);
					this.txtDiasCredito4.Value = (int)drCredCliente.GetInt32(3);
					this.txtDiasCredito5.Value = (int)drCredCliente.GetInt32(4);
				}
				drCredCliente.Close();
			}
			else
			{
				this.txtDiasCredito1.Value = 0;
				this.txtDiasCredito2.Value = 0;
				this.txtDiasCredito3.Value = 0;
				this.txtDiasCredito4.Value = 0;
				this.txtDiasCredito5.Value = 0;

				this.txtDiasCredito1.Enabled = false;
				this.txtDiasCredito2.Enabled = false;
				this.txtDiasCredito3.Enabled = false;
				this.txtDiasCredito4.Enabled = false;
				this.txtDiasCredito5.Enabled = false;
			}
			#endregion Dias de credito				

			if (((int)this.txtIdGrupoCliente.Value == 2 && !bRise) || (int)this.txtIdGrupoCliente.Value == 3)
			{
				this.txtIdTipoDePago.Value = 2;
				this.txtNumero.MaxLength = 20;
			}
			else 
			{
				this.txtIdTipoDePago.Value = 1;
				this.txtNumero.MaxLength = 9;
			}

			#region S.R.I.
			if (((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) && !bRise)
			{
				this.txtSerieFact.Enabled = false;	
				this.txtAutFactura.Enabled = false;	
				this.dtCaducidad.Enabled = false;
				this.dtFechadeEmicion.Enabled = false;
	
				this.txtSerieFact.Text = "";	
				this.txtAutFactura.Text = "";	
				this.dtCaducidad.Value = System.DBNull.Value;
				this.dtFechadeEmicion.Value = DateTime.Today;;
				this.txtNumero.Focus();
			}
			else
			{
				this.txtSerieFact.Enabled = true;	
				this.txtAutFactura.Enabled = true;	
				this.dtCaducidad.Enabled = true;
				this.dtFechadeEmicion.Enabled = true;
				this.txtAutFactura.Focus();
			}
			#endregion S.R.I.			

			#region Liquidacion de Compras
			if((int)this.cmbComprobante.Value == 3)
			{
				bLiquidacion = true;
				this.txtAutFactura.Enabled = false;
				this.txtSerieFact.Enabled = false;
				this.txtNumero.Enabled = false;
				
				if (miAcceso.BLimiteBusquedaFechas)
				{
					this.dtFecha.Enabled = true;
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
					this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);
				}
				else
				{
					this.dtFecha.Enabled = false;
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				}

				this.txtAutFactura.Text = "";
				this.txtSerieFact.Text = "";
				this.txtNumero.Text = "";
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				this.dtCaducidad.Focus();
			}
			else
			{
				bLiquidacion = false;
				this.txtAutFactura.Enabled = true;
				this.txtSerieFact.Enabled = true;
				this.txtNumero.Enabled = true;
				this.dtFecha.Enabled = true;

				this.txtAutFactura.Focus();
			}
			#endregion Liquidacion de Compras
		}


		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Proveedor
				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, (bool)this.chkConsignacion.Checked);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				drC.Read();
				if (!drC.HasRows)
				{
					drC.Close();		
					MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				drC.Close();		
				#endregion Busca Proveedor							

				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", this.txtRuc.Text.ToString())) == 0)
				{
					MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, 6, 1, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtRuc.Text = miBusqueda.Ruc;
						if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) 
						{							
							this.txtNumero.Focus();
						}
						else 
						{
							if(bLiquidacion)
							{
								this.txtAutFactura.Enabled = true;
								this.txtAutFactura.Focus();
							}
							else
							{
								this.txtAutFactura.Focus();
							}
						}
					}
				}
			}

			if (e.KeyCode == Keys.Enter)
			{				
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.txtNumero.Focus();
				else this.txtAutFactura.Focus();
			}
		}

		private void txtSerieFact_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
			if (this.cmbBodega.ActiveRow == null) return;

			this.cmbProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
		}

		private void txtSerieFact_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerieFact_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtSerieFact.Text.ToString().Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					e.Cancel = true;					
				}
				else if (this.txtSerieFact.Text.Length < 6 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3) Funcion.CamposNumericos(sender, e);
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (MenuLatinium.stDirFacturacion != "POINT")
			{
				if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3) 
				{
					char Pad = '0';
					this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);
				}
			}
		}

		private void txtAutFactura_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{				
				if (this.txtAutFactura.Text.ToString().Trim().Length > 0)
				{
					this.txtAutFactura.Text = this.txtAutFactura.Text.Replace(" ", String.Empty);
				}
				if (this.txtAutFactura.Text.Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					e.Cancel = true;					
				}
				else if ((this.txtAutFactura.Text.Length != 10 && this.txtAutFactura.Text.Length != 37 && this.txtAutFactura.Text.Length != 49) && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Longitud de la Autorización debe tener 10, 37 o 49 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void txtAutFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtAutFactura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDiasCredito1_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito1.SelectAll();
		}

		private void txtDiasCredito1_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito1.SelectAll();		
		}

		private void txtDiasCredito2_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito2.SelectAll();
		}

		private void txtDiasCredito2_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito2.SelectAll();
		}

		private void txtDiasCredito3_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito3.SelectAll();
		}

		private void txtDiasCredito3_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito3.SelectAll();
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;
				}
				
				else
				{
					if (MenuLatinium.stDirFacturacion != "POINT")
					{
						char Pad = '0';
						this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);
					}
					
					if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33 || MenuLatinium.stIdDB == 3)
					{
						int iValComprobante = 0;
						iValComprobante = (int)this.cmbComprobante.Value;
						bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeracionCompra '{0}', '{1}', '{2}', {3}, {4}, 4, {5}", 
							this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, iValComprobante), true);
						if (bRes)
						{
							MessageBox.Show("El Número de Documento ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtNumero.Focus();
							e.Cancel = true;					
						}
					}
					else
					{
						int iComprobante = 0;
						iComprobante = (int)this.cmbComprobante.Value;
						bool bResOP = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeracionCompra '{0}', '{1}', '{2}', {3}, {4}, 4, {5}", 
							this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, iComprobante), true);
						if (bResOP)
						{
							MessageBox.Show("El Número de Documento ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtNumero.Focus();
							e.Cancel = true;					
						}
					}
				}
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;				
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbImportacionNombre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbImportacionNombre_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (!bEdicion) return;
			if (e.Button != MouseButtons.Right) return;
			this.cmbImportacionNombre.Value = System.DBNull.Value;			
			cmbImportacionNombre.Focus();
			cmbImportacionNombre.Refresh();
		}

		private void ActualizaIVA()
		{
			if (this.ultraGrid1.Rows.Count > 0)
			{			
				if (this.dtFecha.Value != System.DBNull.Value)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						string sSQL = string.Format("Exec RetornaIVAClienteArticulo {0}, {1}, '{2}'", (int)this.txtIdGrupoCliente.Value, (int)dr.Cells["idArticulo"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						dr.Cells["Impuesto"].Value = Funcion.decEscalarSQL(cdsSeteoF, sSQL); 					
					}
				}
			}
		}

//		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
//		{
//			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
//															
//			#region Iva
//			if (drSeteoF["CambiarIvaC"] == DBNull.Value || !(bool)drSeteoF["CambiarIvaC"]) e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;			
//			#endregion Iva
//
//			#region Formato numerico
//			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
//			string stMoneda = "#,##0.";
//			string stMascara = "nnn,nnn,nnn.";
//			string stCeros = "0000000000";
//			string stNumero = "nnnnnnnnnn";
//			string stFormato = "";
//			string stInput = "";
//			#endregion Formato numerico
//
//			#region Decimales para el total
//			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
//			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
//			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 
//			e.Layout.Bands[0].Columns["DescuentoArt"].Format = stFormato;
//
//			this.txtDescIvaTotal.FormatString = stFormato;
//			this.txtIvaTotal.FormatString = stFormato;
//			this.txtIva.FormatString = stFormato;
//			this.txtIva0.FormatString = stFormato;
//			this.txtTotal.FormatString = stFormato;
//			this.txtDescIva0Total.FormatString = stFormato;
//						
//			this.txtDescIvaTotal.MaskInput = stInput;
//			this.txtIvaTotal.MaskInput = stInput;
//			this.txtIva.MaskInput = stInput;
//			this.txtIva0.MaskInput = stInput;
//			this.txtTotal.MaskInput = stInput;
//			this.txtDescIva0Total.MaskInput = stInput;
//			#endregion Decimales para el total
//
//			#region Decimales para el precio
//			stFormato = stMoneda + stCeros.Substring(0, 5);// (int) drSeteoF["decPrecio"]);
//			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
//			stInput = stMascara + stNumero.Substring(0, 5);// (int) drSeteoF["decPrecio"]);
//			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;			
//			#endregion Decimales para el precio
//			#region Decimales para el Courier
//			stFormato = stMoneda + stCeros.Substring(0, 4);// (int) drSeteoF["decPrecio"]);
//			e.Layout.Bands[0].Columns["Courier"].Format = stFormato;
//			stInput = stMascara + stNumero.Substring(0, 4);// (int) drSeteoF["decPrecio"]);
//			e.Layout.Bands[0].Columns["Courier"].MaskInput = stInput;			
//			#endregion Decimales para Courier
//			stFormato = stMoneda + stCeros.Substring(0, 4);
//			e.Layout.Bands[0].Columns["Peso"].Format = stFormato;		
//			stInput = stMascara + stNumero.Substring(0, 4);
//			e.Layout.Bands[0].Columns["Peso"].MaskInput = stInput;			
//			stFormato = stMoneda + stCeros.Substring(0, 4);
//			e.Layout.Bands[0].Columns["Arancel"].Format = stFormato;		
//			stInput = stMascara + stNumero.Substring(0, 4);
//			e.Layout.Bands[0].Columns["Arancel"].MaskInput = stInput;
//			#region Decimales para el Impuesto
//			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decImpuesto"]);
//			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
//			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decImpuesto"]);
//			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;
//			stInput = stMascara + stNumero.Substring(0, 2);
//			e.Layout.Bands[0].Columns["Valores"].MaskInput = stInput;
//			#endregion Decimales para el Impuesto
//			#region Decimales para porcentaje de descuento
//			stFormato = stMoneda + stCeros.Substring(0, 3);//(int) drSeteoF["decPorcDesc"]);
//			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
//			e.Layout.Bands[0].Columns["DescNC"].Format = stFormato;
//			int iDecPorD = 0;
//			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];
//
//			stInput = stMascara + stNumero.Substring(0, iDecPorD);
//			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
//			e.Layout.Bands[0].Columns["DescNC"].MaskInput = stInput;
//			#endregion Decimales para porcentaje de descuento
//		}

		private void ultraGrid1_InitializeLayout(
			object sender,
			Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e
			)
		{
			if (e.Layout.Bands.Count == 0)
				return;

			UltraGridBand oBand = e.Layout.Bands[0];

			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico

			CultureInfo cultureENUS =
				CultureInfo.CreateSpecificCulture("ja-jp");

			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";

			#endregion

			#region IVA

			if (oBand.Columns.Exists("Impuesto"))
			{
				if (
					drSeteoF["CambiarIvaC"] == DBNull.Value ||
					!(bool)drSeteoF["CambiarIvaC"]
					)
				{
					oBand.Columns["Impuesto"].CellActivation =
						Activation.Disabled;
				}
			}

			#endregion

			#region Decimales total

			stFormato =
				stMoneda +
				stCeros.Substring(
				0,
				(int)drSeteoF["decTotal"]
				);

			stInput =
				stMascara +
				stNumero.Substring(
				0,
				(int)drSeteoF["decTotal"]
				);

			if (oBand.Columns.Exists("SubTotal"))
				oBand.Columns["SubTotal"].Format = stFormato;

			if (oBand.Columns.Exists("DescuentoArt"))
				oBand.Columns["DescuentoArt"].Format = stFormato;

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

			#region Precio

			stFormato =
				stMoneda +
				stCeros.Substring(0, 5);

			stInput =
				stMascara +
				stNumero.Substring(0, 5);

			if (oBand.Columns.Exists("Precio"))
			{
				oBand.Columns["Precio"].Format = stFormato;
				oBand.Columns["Precio"].MaskInput = stInput;
			}

			#endregion

			#region Courier

			stFormato =
				stMoneda +
				stCeros.Substring(0, 4);

			stInput =
				stMascara +
				stNumero.Substring(0, 4);

			if (oBand.Columns.Exists("Courier"))
			{
				oBand.Columns["Courier"].Format = stFormato;
				oBand.Columns["Courier"].MaskInput = stInput;
			}

			#endregion

			#region Peso

			stFormato =
				stMoneda +
				stCeros.Substring(0, 4);

			stInput =
				stMascara +
				stNumero.Substring(0, 4);

			if (oBand.Columns.Exists("Peso"))
			{
				oBand.Columns["Peso"].Format = stFormato;
				oBand.Columns["Peso"].MaskInput = stInput;
			}

			#endregion

			#region Arancel

			stFormato =
				stMoneda +
				stCeros.Substring(0, 4);

			stInput =
				stMascara +
				stNumero.Substring(0, 4);

			if (oBand.Columns.Exists("Arancel"))
			{
				oBand.Columns["Arancel"].Format = stFormato;
				oBand.Columns["Arancel"].MaskInput = stInput;
			}

			#endregion

			#region Impuesto

			stFormato =
				stMoneda +
				stCeros.Substring(
				0,
				(int)drSeteoF["decImpuesto"]
				);

			stInput =
				stMascara +
				stNumero.Substring(
				0,
				(int)drSeteoF["decImpuesto"]
				);

			if (oBand.Columns.Exists("Impuesto"))
			{
				oBand.Columns["Impuesto"].Format = stFormato;
				oBand.Columns["Impuesto"].MaskInput = stInput;
			}

			stInput =
				stMascara +
				stNumero.Substring(0, 2);

			if (oBand.Columns.Exists("Valores"))
			{
				oBand.Columns["Valores"].MaskInput = stInput;
			}

			#endregion

			#region Descuento %

			stFormato =
				stMoneda +
				stCeros.Substring(0, 3);

			if (oBand.Columns.Exists("DescuentoPorc"))
				oBand.Columns["DescuentoPorc"].Format = stFormato;

			if (oBand.Columns.Exists("DescNC"))
				oBand.Columns["DescNC"].Format = stFormato;

			int iDecPorD = 0;

			if (drSeteoF["decPorcDesc"] != null)
				iDecPorD = (int)drSeteoF["decPorcDesc"];

			stInput =
				stMascara +
				stNumero.Substring(0, iDecPorD);

			if (oBand.Columns.Exists("DescuentoPorc"))
				oBand.Columns["DescuentoPorc"].MaskInput = stInput;

			if (oBand.Columns.Exists("DescNC"))
				oBand.Columns["DescNC"].MaskInput = stInput;

			#endregion
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetCompra"].Value = 0;
				e.Row.Cells["idOrigen"].Value = 0;
				e.Row.Cells["Precio"].Value = 0.00m;
				e.Row.Cells["Impuesto"].Value = 0.00m;
				e.Row.Cells["DescuentoArt"].Value = 0.00m;
				e.Row.Cells["DescuentoPorc"].Value = 0.00m;
				e.Row.Cells["DescNC"].Value = 0.00m;
				e.Row.Cells["SubTotal"].Value = 0.00m;
				e.Row.Cells["Bodega"].Value = (int) this.cmbBodega.Value;
				e.Row.Cells["Signo"].Value = IdSigno;
				e.Row.Cells["Tipo"].Value = 0;
				e.Row.Cells["Courier"].Value = 0.00m;
				e.Row.Cells["Peso"].Value = 0.00m;								
				e.Row.Cells["Arancel"].Value = 0.00m;
				e.Row.Cells["Valores"].Value = 0.00m;
				e.Row.Cells["idTipoCalculo"].Value = 0;
				e.Row.Cells["idTipoGasto"].Value = 0;
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if ((int)e.Rows[i].Cells["idOrigen"].Value > 0)
					{					
						MessageBox.Show("No puede Eliminar Filas cuando la Factura tiene Origen en Orden de Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}

					if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;				
					else 
					{						
						if (!bNuevo)
						{
							string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLBorraLinea, true);
						}
						e.DisplayPromptMsg = false;
					}
				}
			}
			else e.DisplayPromptMsg = false;
		}

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			//			if (e.MaskErrorInfo == null)
			//			{
			//				e.Cancel = true;
			//				string stError = e.DataErrorInfo.InvalidValue.ToString();
			//				string stCol = e.DataErrorInfo.Cell.Column.ToString();
			//				string stMensaje = "Error Grilla: Compra";
			//				int iLong = stMensaje.Length;
			//				if (iLong > 20) iLong = 20;
			//
			//				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
			//				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			//
			//				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol
			//					+ "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", 
			//					"Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (this.ultraGrid1.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{
						#region F3
					case (int) Keys.F3:						
						#region Servicios
						if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO") && bEdicion)
						{	
							if ((int)this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value == 0)
							{
								using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString(), 4))
								{
									if (DialogResult.OK == miBusqueda.ShowDialog())
									{										
										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
									}
								}
							}
						}									
						#endregion Servicios
											
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
				MessageBox.Show(ex.Message, "Error en el Grid");
			}
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			//			int iTotalLineas = this.ultraGrid1.Rows.Count;
			//			string stExec = string.Format("Exec CompraNumeroLineas 4, {0}", iTotalLineas);
			//			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);
			//
			//			if (stMensaje.Length > 0) 
			//			{
			//				e.Cancel = true;
			//				MessageBox.Show(stMensaje); // "No puede crear más lineas");
			//			}
		}

		private void cmdTotal()
		{
			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;
	
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;
			this.txtIva0.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				//dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					
					dr.Cells["Subtotal"].Value = Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescNC"].Value);

					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Subtotal"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2);
					//Math.Round(((Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value)) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2);
					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{						
						this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);						
						this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);						
						dIvaLineal = dIvaLineal + 
							((Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value)) * Convert.ToDecimal(dr.Cells["Impuesto"].Value)) / 100.00m;
					}
					else
					{				
						this.txtIva0.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
						this.txtDescIva0Total.Value = Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
				}
				//dPorcIva = 8.00m;
				//this.txtIvaTotal.Value = Math.Round((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva /100, 2);
				this.txtIvaTotal.Value = Math.Round(dIvaLineal, 2);
				//this.txtIvaTotal.Value = dIvaLineal;

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

					using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(stCod, 4))
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
				stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 4, '{2}', 1", IdArt, (int)this.txtIdGrupoCliente.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
				dr.Read();
				e.Cell.Row.Cells["Codigo"].Value = dr.GetString(0);
				e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
				e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
				e.Cell.Row.Cells["Cantidad"].Value = 1;
				//				e.Cell.Row.Cells["Precio"].Value = 0.00;
				e.Cell.Row.Cells["Impuesto"].Value = dr.GetDecimal(4);
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				dr.Close();

				if (MenuLatinium.stIdDB == 1) 
				{
					if (IdArt == 9867)
					{
						Infragistics.Win.UltraWinGrid.UltraGridRow fila;
						fila = e.Cell.Row;
						if (fila.Band.Columns.Exists("idProyecto"))
						{
							fila.Cells["idProyecto"].Value = 93;
						}
					}
					else if (IdArt == 8944)
					{
						Infragistics.Win.UltraWinGrid.UltraGridRow fila;
						fila = e.Cell.Row;
						if (fila.Band.Columns.Exists("idProyecto"))
						{
							fila.Cells["idProyecto"].Value = 87;
						}
					}
				}

				cmdTotal();
        
				VerificaCambio = 2; // Verificaciones y Cambio de precio
				#endregion Asignaciones
			}
			#endregion Codigo o articulo
		
			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["Cantidad"].Value > 0)
					{
						e.Cell.Value = 0.00;
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
						e.Cell.Value = 0.00;
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
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Ingrese un Producto antes de Ingresar el I.V.A.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				if ((decimal)e.Cell.Row.Cells["Impuesto"].Value != dPorcIva && (decimal)e.Cell.Row.Cells["Impuesto"].Value != 0 && (decimal)e.Cell.Row.Cells["Impuesto"].Value != 5.00m && (decimal)e.Cell.Row.Cells["Impuesto"].Value != 8.00m)
				{
					e.Cell.Row.Cells["Impuesto"].Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("El Porcentaje de IVA debe ser 0%, 5%, 8% o {0}%", dPorcIva), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}

				VerificaCambio = 1; // Verificaciones y cambio de precio conficional

				cmdTotal();
			}
			#endregion IVA
			#endregion Fin de Verificacion de Cambio

			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{							
					if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 0.00m)
					{
						e.Cell.Value = 0.00m;
						MessageBox.Show("Ingrese un Producto antes de Ingresar el Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}

				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 100.00m)		
				{
					e.Cell.Row.Cells["DescuentoPorc"].Value = 0.00m;
					MessageBox.Show("El Porcentaje de Descuento no puede ser Mayor a 100", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}

				cmdTotal();
			}

			if (e.Cell.Column.ToString() == "DescuentoArt")
			{				
				cmdTotal();
			}
			#endregion DescuentoPorc o DescuentoArt

			if (e.Cell.Column.ToString() == "idImportacion")
			{
				if (this.cmbImportacion.ActiveRow != null)
				{
					e.Cell.Row.Cells["idTipoCalculo"].Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoCalculo From ImportacionTipo Where idImportacionTipo = {0}", (int)e.Cell.Row.Cells["idImportacion"].Value));
				}
			}
			
			bActualiza = true;
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void GuardarRegistro()
		{
			

			
		}

		private void ImprimeRetencion(int idTipo)
		{
			Cursor = Cursors.WaitCursor;
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idRetencion From Retenciones Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{				
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCompra";
				idCompraRet.Value = (int)this.txtNumIdCompra.Value;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("Retenciones.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Abono");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				if (idTipo == 0) miRepor.Show();
				else miRepor.ImprimeRep(1);
			}			
			Cursor = Cursors.Default;
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

		private string LimpiarValorSQL(string valor)
		{
			if (valor == null)
				return "";
			return valor.Replace("'", "");
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
	
			#region Validacion Encabezado
			#region Busca Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, (bool)this.chkConsignacion.Checked);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (!drC.HasRows)
			{
				drC.Close();		
				MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			drC.Close();		
			#endregion Busca Proveedor
			#region E-Mail
			string sEMail = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Email From Cliente Where idCliente = {0} And (Proveedor = 1 Or Exists (Select 1 From web_usuario wu Where wu.idCliente = Cliente.idCliente))", (int)this.txtNumIdCliente.Value));
												 
			#endregion E-Mail
			#region Serie
			if (!bLiquidacion)
			{
				if (this.txtSerieFact.Text.ToString().Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}

				if (this.txtSerieFact.Text.Length < 6 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}
			}
			#endregion Serie
			#region Numero
			if (!bLiquidacion)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
				if (this.txtNumero.Text.ToString().Length <9)
				{
				
					MessageBox.Show("La longitud del número del documento debe ser de 9 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeracionCompra '{0}', '{1}', '{2}', {3}, {4}, 4, {5}", 
					this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, (int)this.cmbComprobante.Value), true);
				if (bRes)
				{
					MessageBox.Show("El Número de Documento ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();				
					return;
				}
			}
			#endregion Numero
			#region Autorizacion
			if (!bLiquidacion)
			{
				if (this.txtAutFactura.Text.Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					return;
				}
				if ((this.txtAutFactura.Text.Length != 10 && this.txtAutFactura.Text.Length != 37 && this.txtAutFactura.Text.Length != 49) && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Longitud de autorización debe tener 10, 37 o 49 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					return;
				}
			}
			#endregion Autorizacion
			#region Fecha
			if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Fecha
			#region Caducidad
			//			if (this.dtCaducidad.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
			//			{
			//				MessageBox.Show("Seleccione la Fecha de Caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.dtCaducidad.Focus();
			//				return;
			//			}
			//			if (MenuLatinium.stDirFacturacion == "DENTALIS")
			//			{
			//				if (this.dtFechaPeriodo.Value == System.DBNull.Value )
			//				{
			//					MessageBox.Show("Seleccione la Fecha del Periodo de la Compra.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.dtFechaPeriodo.Focus();
			//					return;
			//				}
			//			}
			//
			//			if ((int)this.txtIdTipoRuc.Value != 3)
			//			{
			//				DateTime dtValida = (DateTime) this.dtFecha.Value;
			//				dtValida = dtValida.AddDays(1);
			//
			//				if ((DateTime)this.dtCaducidad.Value < dtValida)
			//				{
			//					MessageBox.Show("Fecha de caducidad no valida.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.dtCaducidad.Focus();
			//					return;
			//				}
			//			}
			#endregion Caducidad
			#region FechadeEmicion
			//if (MenuLatinium.stDirFacturacion == "POINT")
			//{
			if (this.dtFechadeEmicion.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
			{
				MessageBox.Show("Seleccione la Fecha de Emisión.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechadeEmicion.Focus();
				return;
			}
			if ((int)this.txtIdTipoRuc.Value != 3 && this.dtFechadeEmicion.Value != System.DBNull.Value)
			{
				DateTime fecEmision = Convert.ToDateTime(this.dtFechadeEmicion.Value);

				if (fecEmision.Date > DateTime.Today)
				{
					MessageBox.Show("La Fecha de Emisión no puede ser mayor al día de hoy.", "Point Technology",
						MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFechadeEmicion.Focus();
					return;
				}
			}
			//}
			#endregion FechadeEmicion
			#region Dias De Credito
			//			if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3)
			//			{
			//				if ((int)this.txtDiasCredito1.Value == 0 && (int)this.txtDiasCredito2.Value == 0 && (int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value == 0)
			//				{
			//					MessageBox.Show("Ingrese los dias de credito para el pago del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito1.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito1.Value == 0 && ((int)this.txtDiasCredito2.Value > 0 || (int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
			//				{
			//					MessageBox.Show("Dias Credito 1 no puede estar en Cero si Dias Credito 2, 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito1.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito2.Value == 0 && ((int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
			//				{
			//					MessageBox.Show("Dias Credito 2 no puede estar en Cero si Dias Credito 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito2.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value > 0)
			//				{
			//					MessageBox.Show("Dias Credito 3 no puede estar en Cero si Dias Credito 4 es mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito3.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito2.Value > 0 && (int)this.txtDiasCredito1.Value > (int)this.txtDiasCredito2.Value)
			//				{
			//					MessageBox.Show("Dias Credito 1 no puede ser mayor a Dias de Credito 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito2.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito3.Value > 0 && (int)this.txtDiasCredito2.Value > (int)this.txtDiasCredito3.Value)
			//				{
			//					MessageBox.Show("Dias Credito 2 no puede ser mayor a Dias de Credito 3", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito3.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito4.Value > 0 && (int)this.txtDiasCredito3.Value > (int)this.txtDiasCredito4.Value)
			//				{
			//					MessageBox.Show("Dias Credito 3 no puede ser mayor a Dias de Credito 4", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito4.Focus();
			//					return;
			//				}
			//			}
			#endregion Dias De Credito
			#region Forma de Pago SRI
			if ((decimal)this.txtTotal.Value > 1000.00m)
			{
				if (this.cmbTipoFormaPago.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una forma de pago.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipoFormaPago.Focus();
					return;
				}
			}
			#endregion Forma de Pago SRI
			#region Importacion
			if ((int)this.txtIdGrupoCliente.Value == 2 && (int)this.txtIdGrupoCliente.Value == 3)
			{
				if (this.cmbImportacionNombre.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una importación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbImportacionNombre.Focus();
					return;
				}
			}
			#endregion Importacion
			#region CreditoTributario
			if (this.cmbCreditoTributario.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Crédito Tributario para este documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCreditoTributario.Focus();
				return;
			}
			#endregion CreditoTributario
			#region Transaccion en Cero
			if ((decimal) this.txtTotal.Value == 0) 
			{
				MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Transaccion en Cero			
			#endregion Validacion Encabezado

			#region Validacion Detalle
			#region Importacion
			if ((int)this.txtIdGrupoCliente.Value == 2 && (int)this.txtIdGrupoCliente.Value == 3)
			{
				if (this.cmbImportacionNombre.ActiveRow != null)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if (dr.Cells["idImportacion"].Value == DBNull.Value)
						{
							MessageBox.Show(string.Format("Seleccione el tipo de importación para : '{0}'", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
				}
			}
			#endregion Importacion
			#region Valida Cantidad - Precio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int IdArt = (int) dr.Cells["idArticulo"].Value;
					decimal dPrecio = (decimal) dr.Cells["Precio"].Value;

					if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value) 
					{
						MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
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
			#region Valida Centro de Costo
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idProyecto"].Value == DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione el Centro de Costo : '{0}'", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Valida Centro de Costo		
			#region Google Nuevo
			//			if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33 && MenuLatinium.stIdDB != 3)
			//			{
			//				if  ((Decimal) this.txtTotal.Value >= 1000)
			//				{
			//					if(!bNuevo)
			//					{
			//						if (sUrlNube.Length < 10 )
			//						{
			//							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//							return;
			//						}
			//					}
			//					else
			//					{
			//						//if (MenuLatinium.stIdDB != 1 || MenuLatinium.stIdDB != 33 )
			//						//	{
			//						if (txtDirectorioarchivo.Length < 10 )
			//						{
			//							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//							return;
			//						}
			//						//	}
			//					}
			//				}
			//			}
			//			else
			//			{
			//				if ((int)this.txtIdGrupoCliente.Value == 1 ||((int)this.txtIdGrupoCliente.Value == 5 && (Decimal)this.txtTotal.Value >= 5000))
			//					//if  ( (Decimal) this.txtTotal.Value > 5000)
			//				{
			//					if(bNuevo)
			//					{
			//						if (txtDirectorioarchivo.Length < 10 )
			//						{
			//							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//							return;
			//						}
			//					}
			////					else
			////					{
			////						//if (MenuLatinium.stIdDB != 1 || MenuLatinium.stIdDB != 33 )
			////						//	{
			////						if (txtDirectorioarchivo.Length < 10 )
			////						{
			////							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			////							return;
			////						}
			////						//	}
			////					}Menu
			//				}
			//			}
			#endregion Google Nuevo

			if (!ValidaEstadoDocumento())
			return; 

			#endregion Validacion Detalle

			int idProyecto = 0;
			if (this.cmbProyecto.ActiveRow != null) idProyecto = (int)this.cmbProyecto.Value;
			//this.ActualizaIVA();
			//	cmdTotal();

			bool bNumeracionAutomatica = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Automatico From CompraNumero Where idTipoFactura = 85"), true); //Numeracion Automatico Liquidacion Compra

			#region Valida Retencion
			bool bRetencion = false;
			//else if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) bRetencion = false;
			if (this.txtRuc.Text.ToString().Trim() == Funcion.sEscalarSQL(cdsSeteo, "Select LTRIM(RTRIM(Ruc)) From Seteo")) bRetencion = false;
			else if (Funcion.iEscalarSQL(cdsSeteoF, "Select idAgenteRetencion From Seteo") == 5) bRetencion = false;
			else if ((int)this.cmbComprobante.Value == 2 || (int)this.cmbComprobante.Value == 12) bRetencion = false;
			else if (idAgenteRetencion != 2 && idSujetoRetencion == 14) bRetencion = false;
			else if ((int)this.cmbTipoFormaPago.Value == 19) bRetencion = false;
			else
			{
				int iArtRenta = 0;
				int iArtIVA = 0;

				bool bRetieneIVA = false;
				if ((decimal)this.txtIvaTotal.Value > 0) 
					bRetieneIVA = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select IsNull(RetieneIVA, 0) From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value), true);
				
				if ((int)this.cmbComprobante.Value == 1 || (int)this.cmbComprobante.Value == 3)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						int ArtRenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select IsNull(idArticuloSRI, 0) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
						if (ArtRenta != 0 && ArtRenta != 61) iArtRenta ++;
					}

					//					if (iArtRenta > 0)
					//					{
					//						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					//						{
					//							int ArtIVA = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select IsNull(idRetIva, 0) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
					//							if (ArtIVA > 0) iArtIVA ++;
					//						}
					//					}
				}

				if (bRetieneIVA || iArtRenta > 0) bRetencion = true;// || iArtIVA > 0
				else bRetencion = false;
			}
			#endregion Valida Retencion

			#region Variables
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			DateTime dtFechaPago = (DateTime)this.dtFecha.Value;
			DateTime dtFechadeEmicion = (DateTime)this.dtFechadeEmicion.Value;
			DateTime dtCaducidad = DateTime.Today;
			DateTime dtFechaPEriodo  = DateTime.Parse("01/01/2000");
			bool bFecCad = false;
			int IdTipoPago = 0;
			int IdImportacionNombre = 0;
			int IdProyectoDetalle = 0;
			int IdImportacionTipo = 0;
			int idOrigen = 0;
			bool bDiasCompra= false;

			if (this.cmbTipoFormaPago.Value != System.DBNull.Value) IdTipoPago = (int)this.cmbTipoFormaPago.Value;
			if (this.cmbImportacionNombre.ActiveRow != null) IdImportacionNombre = (int)this.cmbImportacionNombre.Value;

			//			if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
			//			{
			//				dtCaducidad = (DateTime)this.dtCaducidad.Value;
			//				bFecCad = true;
			//			}
			if (this.dtFechaPeriodo.Value != System.DBNull.Value)dtFechaPEriodo=(DateTime)this.dtFechaPeriodo.Value;
			#endregion Variables
       
			#region validaPresupuesto
			if(MenuLatinium.stDirFacturacion == "POINT" )
			{
				ValidaPresupuestoCuentaGrid();
			}
			#endregion validaPresupuesto

			if (DialogResult.Yes == MessageBox.Show("Antes de guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n COMO PRECIOS CANTIDADES DESCUENTOS Y DEMAS DATOS DEL PROVEEDOR \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				#region Graba Registro
				if (bLiquidacion || (bRetencion && (int)this.cmbFormaPago.Value == 6))
				{
					#region Registro Con Retenciones
					using (frmRetenciones Retenciones = new frmRetenciones(this.ultraGrid1, (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, 
									 this.txtNumero.Text.ToString(), (DateTime)this.dtFecha.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, true, (int)this.cmbComprobante.Value, bLiquidacion))
					{
						if (DialogResult.OK == Retenciones.ShowDialog())
						{

							using (frmDiasCompra DiasCompra = new frmDiasCompra((int)this.txtNumIdCompra.Value,1))
							{
								if (DialogResult.OK == DiasCompra.ShowDialog())
								{
									if (DiasCompra.uGridDetalle.Rows.Count > 0)
									{	
										#region Proceso con retencion
										#region Valida Pagos Realizados
										decimal dPagado = 0;
										decimal dAutorizado = 0;
										decimal dPA = 0;
										string sSQLValPagos = string.Format("Exec ValidaPagosCompras {0}", (int)this.txtNumIdCompra.Value);
										SqlDataReader drVP = Funcion.rEscalarSQL(cdsSeteoF, sSQLValPagos, true);
										drVP.Read();
										if (drVP.HasRows)
										{
											dPagado = drVP.GetDecimal(0);
											dAutorizado = drVP.GetDecimal(1);
										}
										drVP.Close();

										dPA = dPagado + dAutorizado;
										#endregion Valida Pagos Realizados

										if (Math.Round((decimal)this.txtTotal.Value, 2) == 0)// < Math.Round((dPA + (decimal)Retenciones.txtRetenido.Value), 2))
										{
											MessageBox.Show(string.Format("El Valor Pagado {0} más el Valor Autorizado {1} más el Valor Retenido {2} es Mayor al Valor Total de la Factura {3}", dPagado, dAutorizado, (decimal)Retenciones.txtRetenido.Value, (decimal)this.txtTotal.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										}
										else
										{
											#region Caso Contrario
											using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

													#region Proceso

													#region Variables Liquidacion Compra
													if (bLiquidacion)
													{
														if (bNuevo && bNumeracionAutomatica)
														{
															string sSQLNumero = string.Format("Exec NumeracionLocales 85, 0, 0");
															oCmdActualiza.CommandText = sSQLNumero;
															this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
														}

														string sSQLSerie = string.Format("Select SerieLiquidacion From Seteo");
														oCmdActualiza.CommandText = sSQLSerie;
														this.txtSerieFact.Text = oCmdActualiza.ExecuteScalar().ToString();
													}
													#endregion Variables Liquidacion Compra

													#region Graba Registro
													#region Graba Maestro
													int idProyectoZOOM = 0;
													if (this.cmbProyectoZOOM.ActiveRow != null) idProyectoZOOM = (int)this.cmbProyectoZOOM.Value;
													//
													//													string sSQLMaestro = string.Format("Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}',{39},'{40}'",
													//														(int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtNumidOrigen1.Value,
													//														(int)this.txtNumIdAsiento.Value, (int)this.txtIdSucursal.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbComprobante.Value, 
													//														idProyectoZOOM, (int)this.txtNumContadoCredito.Value, (int)this.txtNumEstado.Value, (int)this.txtNumIdCliente.Value, 
													//														(int)this.cmbBodega.Value, idProyecto, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(),
													//														this.txtAutFactura.Text.ToString(), dtFecha.ToString("yyyyMMdd"), dtCaducidad.ToString("yyyyMMdd"),
													//														(int)this.cmbFormaPago.Value, (int)this.txtDiasCredito1.Value, (int)this.txtDiasCredito2.Value, (int)this.txtDiasCredito3.Value, IdTipoPago,
													//														IdImportacionNombre, this.txtComprobante.Text.ToString(), (bool)this.chkConsignacion.Checked, this.txtNotas.Text.ToString(),
													//														(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
													//														bFecCad, (int)this.txtIdTipoDePago.Value, (int)this.cmbCreditoTributario.Value, dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), (int)this.txtDiasCredito4.Value ,dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),(int)this.txtDiasCredito5.Value,dtFechadeEmicion.ToString("yyyyMMdd HH:mm"));			
													
													string sSQLMaestro = string.Format(
														"Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39}, '{40}'",

														(int)this.txtNumIdCompra.Value,
														(int)this.txtNumidOrigen.Value,
														(int)this.txtNumidOrigen1.Value,
														(int)this.txtNumIdAsiento.Value,
														(int)this.txtIdSucursal.Value,
														(int)this.txtIdTipoFactura.Value,
														(int)this.cmbComprobante.Value,
														idProyectoZOOM,
														(int)this.txtNumContadoCredito.Value,
														(int)this.txtNumEstado.Value,
														(int)this.txtNumIdCliente.Value,
														(int)this.cmbBodega.Value,
														idProyecto,

														LimpiarValorSQL(this.txtSerieFact.Text),
														LimpiarValorSQL(this.txtNumero.Text),
														LimpiarValorSQL(this.txtAutFactura.Text),

														dtFecha.ToString("yyyyMMdd"),
														dtCaducidad.ToString("yyyyMMdd"),

														(int)this.cmbFormaPago.Value,
														(int)this.txtDiasCredito1.Value,
														(int)this.txtDiasCredito2.Value,
														(int)this.txtDiasCredito3.Value,
														IdTipoPago,
														IdImportacionNombre,

														LimpiarValorSQL(this.txtComprobante.Text),

														(bool)this.chkConsignacion.Checked,
														LimpiarValorSQL(this.txtNotas.Text),

														(decimal)this.txtIva0.Value,
														(decimal)this.txtDescIva0Total.Value,
														(decimal)this.txtIva.Value,
														(decimal)this.txtDescIvaTotal.Value,
														(decimal)this.txtIvaTotal.Value,
														(decimal)this.txtTotal.Value,

														bFecCad,
														(int)this.txtIdTipoDePago.Value,
														(int)this.cmbCreditoTributario.Value,

														dtFechaRecepcion.ToString("yyyyMMdd HH:mm"),
														(int)this.txtDiasCredito4.Value,
														dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),
														(int)this.txtDiasCredito5.Value,
														dtFechadeEmicion.ToString("yyyyMMdd HH:mm")
														);
													
													oCmdActualiza.CommandText = sSQLMaestro;
													this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();//Funcion.iEscalarSQL(cdsSeteoF, sSQLMaestro, true);
													#endregion Graba Maestro

													#region Graba Detalle
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
													{	
														if (dr.Cells["idProyecto"].Value != System.DBNull.Value) IdProyectoDetalle = (int)dr.Cells["idProyecto"].Value; else IdProyectoDetalle = 0;
														if (dr.Cells["idImportacion"].Value != System.DBNull.Value) IdImportacionTipo = (int)dr.Cells["idImportacion"].Value; else IdImportacionTipo = 0;
														if (dr.Cells["idOrigen"].Value != System.DBNull.Value) idOrigen = (int)dr.Cells["idOrigen"].Value; else idOrigen = 0;
				
														string sSQLDetalle = string.Format("Exec GrabaDetalleCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}', '{23}'", 
															(int)dr.Cells["idDetCompra"].Value, idOrigen, (int)this.txtNumIdCompra.Value, 
															(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, Math.Round((decimal)dr.Cells["DescuentoArt"].Value, 2), 
															(decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value, (decimal)dr.Cells["DescNC"].Value,
															IdProyectoDetalle, IdImportacionTipo, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, 
															(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, (decimal)dr.Cells["Valores"].Value,
															(int)dr.Cells["idTipoGasto"].Value, dr.Cells["Articulo"].Value, dr.Cells["Codigo"].Value.ToString());
			
														oCmdActualiza.CommandText = sSQLDetalle;
														oCmdActualiza.ExecuteNonQuery();
													}
													#endregion
													
													#region Venta Verde
													MarcarVentaVerdeLiquidada(oCmdActualiza);
													#endregion Venta Verde

													#region Consulta Detalle
													string sSQLActDetalle = string.Format("Exec ConsultaDetalleCompra {0}", (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLActDetalle;
													SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
													DataTable oDTDetalle = new DataTable();
													oDADetalle.Fill(oDTDetalle);
													this.ultraGrid1.DataSource = oDTDetalle;
													#endregion

													#endregion Graba Registro		

													#region valida dias
													if (Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count(*) From CompraDiasPagos Where idCompra = {0}",(int)this.txtNumIdCompra.Value)) != DiasCompra.uGridDetalle.Rows.Count)
													{
														string sSQLBorraDias = string.Format("Exec BorraCompraDiasPagos {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLBorraDias;
														oCmdActualiza.ExecuteNonQuery();
														bDiasCompra=true;
													}	
													#endregion valida dias
													/*1*/
													#region CompraDiasPagos
													if(bDiasCompra)
													{
														foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in DiasCompra.uGridDetalle.Rows.All)
														{
															int iPago = (int)dr.Cells["Pagos"].Value ;	
															int iDias = (int)dr.Cells["Dias"].Value;
															try
															{	

																dtFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT  DATEADD(DAY,{0},'{1}')",iDias,Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));
											
																string sSQLDiasPagos = string.Format("Exec GuardaCompraDiasPagos {0}, {1}, {2},'{3}','{4}'",(int)this.txtNumIdCompra.Value,iPago,iDias,dtFecha.ToString("yyyyMMdd"),dtFechaPago.ToString("yyyyMMdd"));
																oCmdActualiza.CommandText = sSQLDiasPagos;
																oCmdActualiza.ExecuteNonQuery();	
															}
															catch(Exception Exc)
															{								
																MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
															}
														}
													}
													#endregion CompraDiasPagos
								
													#region Autoricacion Factura Liquidacion Compra
													if(bLiquidacion)
													{
														if (bNuevo)
														{
															if (Funcion.bEscalarSQL(cdsSeteoF, "Select FacturaElectronica From SeteoF", false))
															{
																oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 8", (int)this.txtNumIdCompra.Value);;
																string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

																string cadInvertida = invertirCadena(sClaveAcceso.Substring(0, 48));
																string ValM11 = Modulo11(cadInvertida).ToString().Trim();
																string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

																oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, (int)this.txtNumIdCompra.Value);
																oCmdActualiza.ExecuteNonQuery();
															}
														}
													}
													#endregion Autoricacion Factura Liquidacion Compra

													#region Historial Liquidacion Compra
													if(bLiquidacion)
													{
														if (bNuevo)
														{
															string sSQLHLC = string.Format("Insert Into LiquidacionCompra(idCompra) Values ({0})", (int)this.txtNumIdCompra.Value);
															oCmdActualiza.CommandText = sSQLHLC;
															oCmdActualiza.ExecuteNonQuery();
														}
													}
													#endregion Historial Liquidacion Compra

													#region Guarda Retenciones
													DateTime dtFechaRet = (DateTime)Retenciones.dtFecha.Value;
													DateTime dtFechaCad = DateTime.Parse("01/01/2000");
													if (Retenciones.dtCaducidad.Value != System.DBNull.Value) dtFechaCad = (DateTime)Retenciones.dtCaducidad.Value;

													string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 1, 0, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
														(int) Retenciones.txtIdRetencion.Value, (int)this.txtNumIdCompra.Value, Retenciones.txtSerieFact.Text.ToString(), Retenciones.txtNumero.Text.ToString(), Retenciones.txtAutFactura.Text.ToString(), 
														dtFechaCad.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), Retenciones.txtNotas.Text.ToString(), (decimal)Retenciones.txtRetenidoIVA.Value, (decimal)Retenciones.txtRetenidoRenta.Value, 
														(decimal)Retenciones.txtRetenido.Value);
													oCmdActualiza.CommandText = sSQLGR;
													int idRetencion = (int)oCmdActualiza.ExecuteScalar();// Funcion.iEscalarSQL(cdsSeteoF, sSQLGR);

													if (!bNuevo)
													{
														string sSQLValida = string.Format("Exec ValidaDetalleDeRetencion {0}", idRetencion);
														oCmdActualiza.CommandText = sSQLValida;
														oCmdActualiza.ExecuteNonQuery();// Funcion.EjecutaSQL(cdsSeteoF, sSQLValida);
													}

													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Retenciones.ultraGrid1.Rows.All)
													{
														string sSQLGDR = string.Format("Exec GuardaDetalleRetenciones {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
															(int)dr.Cells["idDetalleRetencion"].Value, idRetencion, (bool)dr.Cells["IVARenta"].Value, (int)dr.Cells["idRetencionIVARenta"].Value, (decimal)dr.Cells["Base"].Value, (decimal)dr.Cells["Valor"].Value,
															(decimal)dr.Cells["Porcentaje"].Value);
														oCmdActualiza.CommandText = sSQLGDR;
														oCmdActualiza.ExecuteNonQuery();// Funcion.EjecutaSQL(cdsSeteoF, sSQLGDR, true);
													}

													string sSQLActRetenido = string.Format("Update Compra Set Retenido = {0} Where idCompra = {1}", (decimal)Retenciones.txtRetenido.Value, (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLActRetenido;
													oCmdActualiza.ExecuteNonQuery();// Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Retenido = {0} Where idCompra = {1}", (decimal)Retenciones.txtRetenido.Value, (int)this.txtNumIdCompra.Value));
													#endregion Guarda Retenciones									

													#region Genera pagos
													if (!(bool)this.chkConsignacion.Checked)// && (int)this.cmbFormaPago.Value == 6)//(int)this.txtIdGrupoCliente.Value != 3 && 
													{ 					
														string sSQLPagos = string.Format("Exec GuardaAutorizacionesPagoCompra {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLPagos;
														oCmdActualiza.ExecuteNonQuery();
													}
													#endregion Genera pagos

													#region Saldo
													if ((int)this.cmbFormaPago.Value == 6)
													{
														string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 4, 6", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLSaldo;
														oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo);
													}
													#endregion Saldo

													#region Inventario
													oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 4", (int)this.txtNumIdCompra.Value);
													oCmdActualiza.ExecuteNonQuery();
													#endregion Inventario

													#region Seriales
													ProcesarMovimientoSerialFacturacion(oCmdActualiza);
													#endregion
													

													#region Ordenes De Compra a Compra
													if ((int)this.txtNumidOrigen.Value > 0)
													{

														#region ODC Estado Facturacion
														ProcesarODCEstadoFacturacion(
															oCmdActualiza,
															iEstadoOrigen
															);
														#endregion

														string sSQLActODCC = string.Format("Exec ConsultaODCCompra {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLActODCC;
														SqlDataAdapter oDADetalle1 = new SqlDataAdapter(oCmdActualiza);
														DataTable oDTDetalle1 = new DataTable();
														oDADetalle1.Fill(oDTDetalle1);
														this.ultraGrid2.DataSource = oDTDetalle1;

													}
													#endregion Ordenes De Compra a Compra

													if (bNuevo)
													{
														#region Actualizacion de Costos
														string sSQLAC = string.Format("Exec ActualizaCostosArticulo {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLAC;
														oCmdActualiza.ExecuteNonQuery();
														#endregion Actualizacion de Costos

														#region Actualizacion de precios
														//											string stExec = string.Format("Exec ArticuloActualizaPrecio {0}, 1", (int)this.txtNumIdCompra.Value);
														//											oCmdActualiza.CommandText = stExec;
														//											oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, stExec, true);										
														#endregion Actualizacion de precios
													}

													#region Inserta Historial de Costos
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
													{	
														decimal dCosto = ((Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToInt32(dr.Cells["Cantidad"].Value)) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value))/Convert.ToInt32(dr.Cells["Cantidad"].Value);
														string sSQLCosto = string.Format("Exec GrabaUltimoCosto {0}, 0, {1}, {2}, '{3}', {4}, 0",
															(int)dr.Cells["idArticulo"].Value, (int)this.txtNumIdCompra.Value, (int)dr.Cells["idDetCompra"].Value, dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), dCosto);
														oCmdActualiza.CommandText = sSQLCosto;
														oCmdActualiza.ExecuteNonQuery();
													}
													#endregion Inserta Historial de Costos

													#region Actualiza idCompra en Liquidacion de gastos
													//			if (Funcion.bEscalarSQL(cdsSeteoF, "Select CotizaCreaProyecto From SeteoF", true))
													//			{
													//				if (this.cmbProyectoZOOM.Text.ToString().Length > 0)
													//				{
													//					string sSQLActLiq = string.Format("Exec ActualizaidCompraDetLiqGasto {0}, {1}", (int)this.txtNumIdCompra.Value, (int)this.txtNumIdproyectoZoom.Value);
													//					oCmdActualiza.CommandText = sSQLActLiq;
													//					oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLActLiq);
													//				}
													//			}
													#endregion Actualiza idCompra en Liquidacion de gastos 

													#region Asiento
													if (bAsientoAutomatico)
													{
														#region Factura
														//string stProcedimiento = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec CompraAsientoTipo {0}", (int)this.txtNumIdCompra.Value, true));
														string sSQLAsiento = string.Format("Exec CreaAsiento {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLAsiento;
														oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

														string sSQLIdAsiento = string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLIdAsiento;
														this.txtNumIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();//Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true);
														#endregion Factura

														#region Retencion
														string sSQLAsRet = string.Format("Exec CreaAsientoRetencion {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, 0, 0, '{10}'",
															(int)Retenciones.txtIdAsiento.Value, idRetencion, (int)this.txtNumIdCompra.Value, (int)this.cmbFormaPago.Value, this.txtNombre.Text.ToString(), Retenciones.txtNumero.Text.ToString(),
															this.txtNumero.Text.ToString(), dtFechaRet.ToString("yyyyMMdd"), idProyecto, (decimal)Retenciones.txtRetenido.Value, this.txtRuc.Text.ToString());
														oCmdActualiza.CommandText = sSQLAsRet;
														oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLAsRet, true);
														#endregion Retencion
													}
													#endregion Asiento

													#region Valida Estado de Importacion FOB
													if ((int)this.txtIdGrupoCliente.Value == 3)
													{ 
														string sSQLImp = string.Format("Exec ActualizaEstadoImportacionFOB {0}, {1}", (int)this.txtNumIdCompra.Value, iEstadoOrigen);
														oCmdActualiza.CommandText = sSQLImp;
														oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLImp, true);
													}
													#endregion Valida Estado de Importacion FOB

													if (!bNuevo)
													{
														string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLBloqueo;
														oCmdActualiza.ExecuteNonQuery();
													}
													else 
													{
														if ((int)this.txtNumidOrigen.Value > 0 && !bProforma)
														{
															string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
															oCmdActualiza.CommandText = sSQLBloqueoOrigen;
															oCmdActualiza.ExecuteNonQuery();
														}

														if ((int)this.txtNumidOrigen1.Value > 0)
														{
															string sSQLBloqueoOrigen1 = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
															oCmdActualiza.CommandText = sSQLBloqueoOrigen1;
															oCmdActualiza.ExecuteNonQuery();
														}
													}

													if (bProforma)
													{
														string sSQLProforma = string.Format("Update Proforma Set Estado = 3, FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME(), idCompra = {1} Where idProforma = {0}", 
															(int)this.txtNumidOrigen.Value, (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLProforma;
														oCmdActualiza.ExecuteNonQuery();
													}

													oTransaction.Commit();

													//													#region Impresión
													//													this.btnImprimir_Click(sender, e); // REPORTE DE COMPRA
													//													ImprimeRetencion(0);
													//													#endregion Impresión

													#region Controles
													bNuevo = false;
													bEdicion = false;

													this.txtRuc.Enabled = false;
													this.txtSerieFact.Enabled = false;
													this.txtNumero.Enabled = false;
													this.txtAutFactura.Enabled = false;
													this.dtCaducidad.Enabled = false;
													this.dtFechaPeriodo.Enabled = false;
													this.dtFechaPeriodo.Enabled = false;
													this.txtDiasCredito1.Enabled = false;
													this.txtDiasCredito2.Enabled = false;
													this.txtDiasCredito3.Enabled = false;
													this.txtDiasCredito4.Enabled = false;
													this.txtDiasCredito5.Enabled = false;
													this.cmbTipoFormaPago.Enabled = false;
													this.cmbProyectoZOOM.Enabled = false;
													this.cmbImportacionNombre.Enabled = false;
													this.cmbCreditoTributario.Enabled = false;
													this.dtFecha.Enabled = false;
													this.cmbFormaPago.Enabled = false;
													this.dtFechadeEmicion.Enabled = false;
										
			
													ModificaGrilla(false);
			
													this.txtNotas.Enabled = false;
						
													if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
													if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
													if (miAcceso.Editar) this.btnEditar.Enabled = true;
													DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
													DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
													if (miAcceso.EditarSoloDia)
													{
														if (dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString() )
														{
															this.btnEditar.Enabled = true;
														}
													}
													if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
													else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;			
													if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
													if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
													if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
													if (miAcceso.Retencion) this.btnRetencion.Enabled = true;
													this.btnGuardar.Enabled = false;
									

													//										if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33 )
													//										{
													if (sUrlNube.Length > 10 )
													{
														this.btnSubir.Visible =false;
														this.linkLabel1.Visible = true;
													}
													//										}

													#endregion Controles						

													#endregion Proceso

													if (bConsignacion) this.DialogResult = DialogResult.OK;
												}
												catch (Exception ex)
												{
													#region Exception
													try
													{
														oTransaction.Rollback();

														MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
													}
													catch (Exception ex2)
													{
														MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
													}
													#endregion Exception
												}
												finally
												{
													oConexion.Close();

													#region PresupuestoA
													EjecutarPostPresupuesto((int)this.txtNumIdCompra.Value);
													#endregion PresupuestoA
											
												}
											}
											#endregion Caso Contrario
										}
										#endregion Proceso con retencion
									}
								}
							}
						}
					}
					#endregion Registro Con Retenciones
				}
				else
				{
					using (frmDiasCompra DiasCompra = new frmDiasCompra((int)this.txtNumIdCompra.Value,1))
					{
						if (DialogResult.OK == DiasCompra.ShowDialog())
						{
							if (DiasCompra.uGridDetalle.Rows.Count > 0)
							{	
								#region CasoContrario
								#region Valida Pagos Realizados
								decimal dPagado = 0.00m;
								decimal dAutorizado = 0.00m;
								decimal dPA = 0.00m;
								string sSQLValPagos = string.Format("Exec ValidaPagosCompras {0}", (int)this.txtNumIdCompra.Value);
								SqlDataReader drVP = Funcion.rEscalarSQL(cdsSeteoF, sSQLValPagos, true);
								drVP.Read();
								if (drVP.HasRows)
								{
									dPagado = drVP.GetDecimal(0);
									dAutorizado = drVP.GetDecimal(1);
								}
								drVP.Close();

								dPA = dPagado + dAutorizado;
								#endregion Valida Pagos Realizados


								decimal total = Math.Round((decimal)this.txtTotal.Value, 2);
								if (total < dPA)
								{
									MessageBox.Show(string.Format("El Valor Pagado {0} más el Valor Autorizado {1} es Mayor al Valor Total de la Factura {2}", dPagado, dAutorizado, (decimal)this.txtTotal.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								}
								else
								{
									using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

											#region Proceso

											#region Graba Registro
											#region Graba Maestro
											int idProyectoZOOM = 0;
											if (this.cmbProyectoZOOM.ActiveRow != null) idProyectoZOOM = (int)this.cmbProyectoZOOM.Value;

											//											string sSQLMaestro = string.Format("Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39},'{40}'",
											//												(int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtNumidOrigen1.Value,
											//												(int)this.txtNumIdAsiento.Value, (int)this.txtIdSucursal.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbComprobante.Value, 
											//												idProyectoZOOM, (int)this.txtNumContadoCredito.Value, (int)this.txtNumEstado.Value, (int)this.txtNumIdCliente.Value, 
											//												(int)this.cmbBodega.Value, idProyecto, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(),
											//												this.txtAutFactura.Text.ToString(), dtFecha.ToString("yyyyMMdd"), dtCaducidad.ToString("yyyyMMdd"),
											//												(int)this.cmbFormaPago.Value, (int)this.txtDiasCredito1.Value, (int)this.txtDiasCredito2.Value, (int)this.txtDiasCredito3.Value, IdTipoPago,
											//												IdImportacionNombre, this.txtComprobante.Text.ToString(), (bool)this.chkConsignacion.Checked, this.txtNotas.Text.ToString(),
											//												(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, total, 
											//												bFecCad, (int)this.txtIdTipoDePago.Value, (int)this.cmbCreditoTributario.Value, 
											//												dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), (int)this.txtDiasCredito3.Value, dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),
											//												(int)this.txtDiasCredito5.Value,dtFechadeEmicion.ToString("yyyyMMdd HH:mm"));

											string sSQLMaestro = string.Format(
												"Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39}, '{40}'",

												(int)this.txtNumIdCompra.Value,
												(int)this.txtNumidOrigen.Value,
												(int)this.txtNumidOrigen1.Value,
												(int)this.txtNumIdAsiento.Value,
												(int)this.txtIdSucursal.Value,
												(int)this.txtIdTipoFactura.Value,
												(int)this.cmbComprobante.Value,
												idProyectoZOOM,
												(int)this.txtNumContadoCredito.Value,
												(int)this.txtNumEstado.Value,
												(int)this.txtNumIdCliente.Value,
												(int)this.cmbBodega.Value,
												idProyecto,

												LimpiarValorSQL(this.txtSerieFact.Text),
												LimpiarValorSQL(this.txtNumero.Text),
												LimpiarValorSQL(this.txtAutFactura.Text),

												dtFecha.ToString("yyyyMMdd"),
												dtCaducidad.ToString("yyyyMMdd"),

												(int)this.cmbFormaPago.Value,
												(int)this.txtDiasCredito1.Value,
												(int)this.txtDiasCredito2.Value,
												(int)this.txtDiasCredito3.Value,
												IdTipoPago,
												IdImportacionNombre,

												LimpiarValorSQL(this.txtComprobante.Text),

												(bool)this.chkConsignacion.Checked,
												LimpiarValorSQL(this.txtNotas.Text),

												(decimal)this.txtIva0.Value,
												(decimal)this.txtDescIva0Total.Value,
												(decimal)this.txtIva.Value,
												(decimal)this.txtDescIvaTotal.Value,
												(decimal)this.txtIvaTotal.Value,
												total,

												bFecCad,
												(int)this.txtIdTipoDePago.Value,
												(int)this.cmbCreditoTributario.Value,

												dtFechaRecepcion.ToString("yyyyMMdd HH:mm"),
												(int)this.txtDiasCredito3.Value,
												dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),
												(int)this.txtDiasCredito5.Value,
												dtFechadeEmicion.ToString("yyyyMMdd HH:mm")
												);

											oCmdActualiza.CommandText = sSQLMaestro;
											this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
											#endregion Graba Maestro

											#region Graba Detalle
											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
											{	
												if (dr.Cells["idProyecto"].Value != System.DBNull.Value) IdProyectoDetalle = (int)dr.Cells["idProyecto"].Value; else IdProyectoDetalle = 0;
												if (dr.Cells["idImportacion"].Value != System.DBNull.Value) IdImportacionTipo = (int)dr.Cells["idImportacion"].Value; else IdImportacionTipo = 0;
									
												string sSQLDetalle = string.Format("Exec GrabaDetalleCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}', '{23}'", 
													(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value, 
													(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, Math.Round((decimal)dr.Cells["DescuentoArt"].Value, 2), (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value, (decimal)dr.Cells["DescNC"].Value,
													IdProyectoDetalle, IdImportacionTipo, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, 
													(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, (decimal)dr.Cells["Valores"].Value,
													(int)dr.Cells["idTipoGasto"].Value, dr.Cells["Articulo"].Value, dr.Cells["Codigo"].Value.ToString());
								
												oCmdActualiza.CommandText = sSQLDetalle;
												oCmdActualiza.ExecuteNonQuery();
											}
											#endregion

											#region Refrescar Detalle
											string sSQLActDetalle = string.Format("Exec ConsultaDetalleCompra {0}", (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLActDetalle;
											SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
											DataTable oDTDetalle = new DataTable();
											oDADetalle.Fill(oDTDetalle);
											this.ultraGrid1.DataSource = oDTDetalle;
											#endregion
											
											#endregion Graba Registro

											#region valida dias
											if (Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count(*) From CompraDiasPagos Where idCompra = {0}",(int)this.txtNumIdCompra.Value)) != DiasCompra.uGridDetalle.Rows.Count)
											{
												string sSQLBorraDias = string.Format("Exec BorraCompraDiasPagos {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLBorraDias;
												oCmdActualiza.ExecuteNonQuery();
												bDiasCompra=true;
											}	
											#endregion valida dias
											/*2*/

											#region CompraDiasPagos
											if(bDiasCompra)
											{
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in DiasCompra.uGridDetalle.Rows.All)
												{
													int iPago = (int)dr.Cells["Pagos"].Value ;	
													int iDias = (int)dr.Cells["Dias"].Value;
													try
													{	

														dtFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT  DATEADD(DAY,{0},'{1}')",iDias,Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));
											
														string sSQLDiasPagos = string.Format("Exec GuardaCompraDiasPagos {0}, {1}, {2},'{3}','{4}'",(int)this.txtNumIdCompra.Value,iPago,iDias,dtFecha.ToString("yyyyMMdd"),dtFechaPago.ToString("yyyyMMdd"));
														oCmdActualiza.CommandText = sSQLDiasPagos;
														oCmdActualiza.ExecuteNonQuery();	
													}
													catch(Exception Exc)
													{								
														MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
													}
												}
											}
											#endregion CompraDiasPagos

											#region Graba Google
													
											//								if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33 )
											//								{
											if(bNuevo)
											{
												if(txtDirectorioarchivo.Length > 10 )
												{
													NuevoGoogle (txtDirectorioarchivo);
												}
											}
											if (sUrlNube.Length > 10 )
											{
												string sSQLGoogle = string.Format("Exec GrabaUrlNube {0}, '{1}', '{2}'", (int)this.txtNumIdCompra.Value, sCarpeta, sUrlNube );
												oCmdActualiza.CommandText = sSQLGoogle;
												oCmdActualiza.ExecuteNonQuery();
											}
											//	}
												
											#endregion Graba Google

											#region Genera pagos 
											if (!(bool)this.chkConsignacion.Checked)// && (int)this.cmbFormaPago.Value == 6)//(int)this.txtIdGrupoCliente.Value != 3 && 
											{ 
												string sSQLPagos = string.Format("Exec GuardaAutorizacionesPagoCompra {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLPagos;
												oCmdActualiza.ExecuteNonQuery();
											}
											#endregion Genera pagos 

											#region Inventario
											oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, 4", (int)this.txtNumIdCompra.Value);
											oCmdActualiza.ExecuteNonQuery();
											#endregion Inventario

											#region Saldo
											string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 4, 6", (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLSaldo;
											oCmdActualiza.ExecuteNonQuery();								
											#endregion Saldo

											#region Seriales
											ProcesarMovimientoSerialFacturacion(oCmdActualiza);
											#endregion

											#region Ordenes De Compra a Compra
											if ((int)this.txtNumidOrigen.Value > 0)
											{
												
												#region ODC Estado Facturacion
												ProcesarODCEstadoFacturacion(												 
													oCmdActualiza,
													iEstadoOrigen
													);
												#endregion
												string sSQLActODCC = string.Format("Exec ConsultaODCCompra {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLActODCC;
												SqlDataAdapter oDADetalle1 = new SqlDataAdapter(oCmdActualiza);
												DataTable oDTDetalle1 = new DataTable();
												oDADetalle1.Fill(oDTDetalle1);
												this.ultraGrid2.DataSource = oDTDetalle1;		
														
											}
											#endregion Ordenes De Compra a Compra

											if (bNuevo)
											{
												#region Actualizacion de precios
												//									string stExec = string.Format("Exec ArticuloActualizaPrecio {0}, 1", (int)this.txtNumIdCompra.Value);
												//									oCmdActualiza.CommandText = stExec;
												//									oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
												#endregion Actualizacion de precios
												#region Actualizacion de Costos
												if ((int)this.txtIdGrupoCliente.Value != 2)
												{
													if ((int)this.txtIdGrupoCliente.Value != 3)
													{
														string sSQLAC = string.Format("Exec ActualizaCostosArticulo {0}", (int)this.txtNumIdCompra.Value);
														oCmdActualiza.CommandText = sSQLAC;
														oCmdActualiza.ExecuteNonQuery();
													}
												}
												#endregion Actualizacion de Costos								
											}

											#region Inserta Historial de Costos
											if ((int)this.txtIdGrupoCliente.Value != 2)
											{
												if ((int)this.txtIdGrupoCliente.Value != 3)
												{
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
													{	
														decimal dCosto = ((Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToInt32(dr.Cells["Cantidad"].Value)) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value))/Convert.ToInt32(dr.Cells["Cantidad"].Value);
														string sSQLCosto = string.Format("Exec GrabaUltimoCosto {0}, 0, {1}, {2}, '{3}', {4}, 0",
															(int)dr.Cells["idArticulo"].Value, (int)this.txtNumIdCompra.Value, (int)dr.Cells["idDetCompra"].Value, dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), dCosto);
														oCmdActualiza.CommandText = sSQLCosto;
														oCmdActualiza.ExecuteNonQuery();
													}
												}
											}
											#endregion Inserta Historial de Costos

											#region Actualiza idCompra en Liquidacion de gastos
											//			if (Funcion.bEscalarSQL(cdsSeteoF, "Select CotizaCreaProyecto From SeteoF", true))
											//			{
											//				if (this.cmbProyectoZOOM.Text.ToString().Length > 0)
											//				{
											//					string sSQLActLiq = string.Format("Exec ActualizaidCompraDetLiqGasto {0}, {1}", (int)this.txtNumIdCompra.Value, (int)this.txtNumIdproyectoZoom.Value);
											//					oCmdActualiza.CommandText = sSQLActLiq;
											//					oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLActLiq);
											//				}
											//			}
											#endregion Actualiza idCompra en Liquidacion de gastos 

											#region Asiento
											if (bAsientoAutomatico)
											{
												//							string stProcedimiento = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec CompraAsientoTipo {0}", (int)this.txtNumIdCompra.Value, true));
												string sSQLAsiento = string.Format("Exec CreaAsiento {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLAsiento;
												oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

												string sSQLIdAsiento = string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLIdAsiento;
												this.txtNumIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();//this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true);
											}
											#endregion Asiento

											#region Estado
											//						int iEstOrigen = 9; 
											//						if (iEstadoOrigen == 1) iEstOrigen = 12;
											//						string sSQLEstado = string.Format("Update Compra Set Estado = {1} Where idCompra = {0}", (int)this.txtNumidOrigen.Value, iEstOrigen);
											//						oCmdActualiza.CommandText = sSQLEstado;
											//						oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLEstado, true);

											if ((int)this.txtNumEstado.Value == 7) this.lblEstado.Text = "POR LIQUIDAR";
											if ((int)this.txtNumEstado.Value == 6) this.lblEstado.Text = "TEMPORAL";
											if ((int)this.txtNumEstado.Value == 8) this.lblEstado.Text = "LIQUIDADO";
											if ((int)this.txtNumEstado.Value == 14) this.lblEstado.Text = "LIQUIDADO";
											#endregion Estado

											#region Valida Estado de Importacion FOB
											if ((int)this.txtIdGrupoCliente.Value == 3)
											{ 
												string sSQLImp = string.Format("Exec ActualizaEstadoImportacionFOB {0}, {1}", (int)this.txtNumIdCompra.Value, iEstadoOrigen);
												oCmdActualiza.CommandText = sSQLImp;
												oCmdActualiza.ExecuteNonQuery();//Funcion.EjecutaSQL(cdsSeteoF, sSQLImp, true);
											}
											#endregion Valida Estado de Importacion FOB

											/////////////

											if (!bNuevo)
											{
												string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLBloqueo;
												oCmdActualiza.ExecuteNonQuery();
											}
											else 
											{
												if ((int)this.txtNumidOrigen.Value > 0 && !bProforma)
												{
													string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
													oCmdActualiza.CommandText = sSQLBloqueoOrigen;
													oCmdActualiza.ExecuteNonQuery();
												}

												if ((int)this.txtNumidOrigen1.Value > 0)
												{
													string sSQLBloqueoOrigen1 = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
													oCmdActualiza.CommandText = sSQLBloqueoOrigen1;
													oCmdActualiza.ExecuteNonQuery();
												}
											}

											if (bProforma)
											{
												string sSQLProforma = string.Format("Update Proforma Set Estado = 3, FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME(), idCompra = {1} Where idProforma = {0}", 
													(int)this.txtNumidOrigen.Value, (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLProforma;
												oCmdActualiza.ExecuteNonQuery();
											}

											oTransaction.Commit();
											//
											//											#region Impresión
											//											this.btnImprimir_Click(sender, e); // REPORTE DE COMPRA
											//											#endregion Impresión

											#region Controles
											bNuevo = false;
											bEdicion = false;

											this.txtRuc.Enabled = false;
											this.txtSerieFact.Enabled = false;
											this.txtNumero.Enabled = false;
											this.txtAutFactura.Enabled = false;
											this.dtCaducidad.Enabled = false;
											this.txtDiasCredito1.Enabled = false;
											this.txtDiasCredito2.Enabled = false;
											this.txtDiasCredito3.Enabled = false;
											this.txtDiasCredito4.Enabled = false;
											this.txtDiasCredito5.Enabled = false;
											//this.cmbFormaPago.Value = false;
											this.cmbTipoFormaPago.Enabled = false;
											this.cmbImportacionNombre.Enabled = false;
											this.cmbProyectoZOOM.Enabled = false;
											this.cmbCreditoTributario.Enabled = false;
											this.dtFecha.Enabled = false;
											this.dtFechadeEmicion.Enabled = false;
											this.cmbFormaPago.Enabled = false;
													
								
											ModificaGrilla(false);
								
											this.txtNotas.Enabled = false;
											
											if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
											if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
											if (miAcceso.Editar) this.btnEditar.Enabled = true;
											DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
											DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
											if (miAcceso.EditarSoloDia)
											{
												if (dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString() )
												{
													this.btnEditar.Enabled = true;
												}
											}
											if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
											else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;			
											if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
											if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
											if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
											if (miAcceso.Retencion) this.btnRetencion.Enabled = true;
											this.btnGuardar.Enabled = false;
											this.btnTemporal.Enabled = false;
													
											this.btnCancelar.Enabled = true;

													
											//								if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33 )
											//								{
											if (sUrlNube.Length > 10 )
											{
												this.btnSubir.Visible =false;
												this.linkLabel1.Visible = true;
											}
											//								}
													

											#endregion Controles						

											#endregion Proceso

											if (bConsignacion) this.DialogResult = DialogResult.OK;
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
											#region PresupuestoB
											EjecutarPostPresupuesto((int)this.txtNumIdCompra.Value);
											#endregion PresupuestoB
										}
									}
								}
								#endregion CasoContrario
							}
						}
					}
				}
				#endregion Graba Registro
			}
		}

		#region Seriales
		private void ProcesarMovimientoSerialFacturacion(
			SqlCommand oCmdActualiza
			)
		{
			#region Seriales

			if (!MenuLatinium.bConsignacionVenta)
			{
				if (bNuevo && (int)this.txtNumidOrigen.Value > 0)
				{
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						string sSQLMovSerial = string.Format(
							"Exec MovimientoSerialFacturacion {0}, {1}, {2}, 0",
							(int)dr.Cells["idOrigen"].Value,
							(int)dr.Cells["idDetCompra"].Value,
							(int)dr.Cells["idArticulo"].Value
							);

						oCmdActualiza.CommandText = sSQLMovSerial;
						oCmdActualiza.ExecuteNonQuery();
					}
				}
			}

			#endregion
		}

		#endregion

		#region Ordenes De Compra a Compra
		private void ProcesarODCEstadoFacturacion(
			SqlCommand oCmdActualiza,
			int iEstadoOrigen
			)
		{
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow drODCC in ultraGrid2.Rows.All)
			{
				string sSQLODCmp = string.Format(
					"EXEC ODCCompra {0}, {1}, {2}",
					(int)drODCC.Cells["idOrdenDeCompraCompra"].Value,
					(int)drODCC.Cells["idOrdenDeCompra"].Value,
					(int)this.txtNumIdCompra.Value
					);

				oCmdActualiza.CommandText = sSQLODCmp;
				oCmdActualiza.ExecuteNonQuery();

				int idOC = (int)drODCC.Cells["idOrdenDeCompra"].Value;

				/* =========================================================
					 LÓGICA LEGACY (MISMA CONEXIÓN - CORREGIDO)
				========================================================= */

				oCmdActualiza.CommandText = string.Format(
					"EXEC ConsultaPendientesFacturacionODC {0}",
					idOC
					);

				int pendientes = Convert.ToInt32(oCmdActualiza.ExecuteScalar());

				/* =========================================================
					 SOLO SI NO HAY PENDIENTES SE ACTUALIZA ESTADO
				========================================================= */
				if (pendientes == 0)
				{
					int iEstOrigen = 9;

					if (iEstadoOrigen == 1)
						iEstOrigen = 12;

					string sSQLEstado = string.Format(
						"UPDATE Compra SET Estado = {1} WHERE idCompra = {0}",
						idOC,
						iEstOrigen
						);

					oCmdActualiza.CommandText = sSQLEstado;
					oCmdActualiza.ExecuteNonQuery();
				}

				
				/* =========================================
								 CONTROL INTERNO FACTURACIÓN SERIAL
							========================================= */
				string sSQLRevertir = string.Format(
					"EXEC OrdenDeCompraRevertirFacturacionODC {0}, {1}, '{2}'",
					idOC,
					iEstadoOrigen,
					MenuLatinium.stUsuario
					);

				oCmdActualiza.CommandText = sSQLRevertir;
				oCmdActualiza.ExecuteNonQuery();

			}


		}
		#endregion

		private void EjecutarPostPresupuesto(int idCompra)
		{
			if (MenuLatinium.stDirFacturacion != "POINT") return;
			if (idCompra <= 0) return;

			int resultado = ValidaPresupuestoCuentaGrid10(); // 1 = OK, 0 = encontró mensajes

			string url = MenuLatinium.stApiPresupuestoMensaje + idCompra;
			Funcion.ConsumirAPISilencioso(url);
		}

		private void dtCaducidad_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtCaducidad.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Seleccione la Fecha de Caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtCaducidad.Focus();
					e.Cancel = true;					
				}
				else
				{
					DateTime dtValida = (DateTime) this.dtFecha.Value;
					dtValida = dtValida.AddDays(1);

					if ((DateTime)this.dtCaducidad.Value < dtValida)
					{
						MessageBox.Show("Fecha de caducidad no valida.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.dtCaducidad.Focus();
						e.Cancel = true;						
					}
				}
			}
		}

		private void dtCaducidad_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{	
			this.Consultar((int)this.txtNumIdCompra.Value);

			#region Liquidacion Compra Validacion SRI
			if (bLiquidacion)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFE_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) == 2)
				{
					MessageBox.Show("No puede Editar la Liquidacion de Compra, esta autorizada por el S.R.I.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Liquidacion Compra Validacion SRI

			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (!miAcceso.Editar && !miAcceso.EditarSoloDia)
			{
				MessageBox.Show("No tiene Acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Editar esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}		
			else if (iOrigen == 1)
			{
				MessageBox.Show("No puede editar una Factura de Caja Chica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (iOrigen == 2)
			{
				MessageBox.Show("No puede editar una Factura de Consumos con Tarjeta de Crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}			
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos Where Estado = 1 And idTipo = 4 And idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("Existen Pagos Autorizados\n\n Desautorice los Pagos o \n\nConsulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (!miAcceso.BEditarPagosRealizados && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 0 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("Existen Pagos Hechos de esta Factura\n\nConsulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if ((int)this.txtNumEstado.Value == 8 && !miAcceso.BEditarLiquidado)
			{
				MessageBox.Show("La Factura Esta Liquidada\n\nConsulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}		
			else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))
			{
				this.txtNumero.Enabled = true;
			
				if (miAcceso.BEditarNombre)
					if ((int)this.txtNumidOrigen.Value == 0) this.txtRuc.Enabled = true;

				if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3)
				{
					this.txtSerieFact.Enabled = true;
					this.txtAutFactura.Enabled = true;
					this.dtCaducidad.Enabled = true;
					this.dtFechaPeriodo.Enabled = true;
				}
				
				if ((int)this.cmbFormaPago.Value == 6)
				{
					if ((int)this.txtNumEstado.Value == 7)
					{
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 0 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) == 0)
						{
							this.txtDiasCredito1.Enabled = false;
							this.txtDiasCredito2.Enabled = false;
							this.txtDiasCredito3.Enabled = false;
							this.txtDiasCredito4.Enabled = false;
							this.txtDiasCredito5.Enabled = false;
						}
					}
				}

				if ((int)this.txtNumidOrigen.Value == 0) this.cmbCreditoTributario.Enabled = true;

				this.cmbTipoFormaPago.Enabled = true;

				if (bGastosPorProyecto) this.cmbProyectoZOOM.Enabled = true; else this.cmbImportacionNombre.Enabled = true;

				if (miAcceso.CambiarFecha) 
				{
					this.dtFecha.Enabled = true;
					this.dtFechadeEmicion.Enabled = true;
				}

				if (miAcceso.EditarDetalleFormaPago) 
				{
					this.cmbFormaPago.Enabled = true;
				}
				else
				{
					this.cmbFormaPago.Enabled = false;
				}
			
				ModificaGrilla(true);
			
				this.txtNotas.Enabled = true;
						
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = true;
				if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
				this.btnAsiento.Enabled = false;
				this.btnRetencion.Enabled = false;
				this.btnBuscarOrdenDeCompra.Enabled = false;
				this.btnProforma.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;
				if (sUrlNube.Length > 10 )
				{
					this.btnSubir.Visible =false;
					this.linkLabel1.Visible = true;
				}
				#region Edicion Liquidacion Compra
				if(bLiquidacion)
				{
					this.txtAutFactura.Enabled = false;
					this.txtSerieFact.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtNumero.Enabled = false;
				}
				#endregion Edicion Liquidacion Compra

				//if (sUrlNube.Length < 10 )
				//this.btnSubir.Visible = true;

				//if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33)				

				if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3) this.txtAutFactura.Focus(); else this.txtNumero.Focus();			

				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
				
				bNuevo = false;
				bEdicion = true;
			}
		}

		private void btnBuscarOrdenDeCompra_Click(object sender, System.EventArgs e)
		{
			if (!bEdicion)
			{
				this.btnCancelar_Click(sender, e);
				string sIdCompra = "";

				using (BuscaOrdenCompra miBusqueda = new BuscaOrdenCompra())
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						try
						{
							int iCont = 0;
							int IBodega = 0;

							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in miBusqueda.grdOrdenCompra.Rows.All)
							{
								if ((bool)dr.Cells["Sel"].Value)
								{
									if (iCont == 0)
									{
										#region datos
										this.txtNumidOrigen.Value = (int)dr.Cells["idCompra"].Value;
										this.chkConsignacion.Checked = (bool)dr.Cells["Consignacion"].Value;
										iEstadoOrigen = (int)dr.Cells["Estado"].Value;
										this.txtNumContadoCredito.Value = (int)dr.Cells["ContadoCredito"].Value;
										this.txtNumIdCliente.Value = (int)dr.Cells["idCliente"].Value;
										this.txtIdTipoRuc.Value = (int)dr.Cells["idTipoRuc"].Value;
										this.txtIdGrupoCliente.Value = (int)dr.Cells["idGrupoCliente"].Value;
										// this.txtIdTipoDePago.Value = (int)dr.Cells["idTipoDePago"].Value;
										this.cmbComprobante.Value = (int)dr.Cells["idComprobante"].Value;
										this.txtRuc.Text = dr.Cells["Ruc"].Value.ToString();
										this.txtNombre.Text = dr.Cells["Nombre"].Value.ToString();
										IBodega = (int)dr.Cells["Bodega"].Value;
										this.txtDiasCredito1.Value = (int)dr.Cells["DiasCredito"].Value;
										this.txtDiasCredito2.Value = (int)dr.Cells["DiasCredito1"].Value;
										this.txtDiasCredito3.Value = (int)dr.Cells["DiasCredito2"].Value;

										if (dr.Cells["idImportacionTipo"].Value != System.DBNull.Value)
											this.cmbImportacionNombre.Value = (int)dr.Cells["idImportacionTipo"].Value;

										this.txtComprobante.Text = dr.Cells["Numero"].Value.ToString();
										this.txtNotas.Text = dr.Cells["Notas"].Value.ToString();
										this.txtIva0.Value = (decimal)dr.Cells["SubtotalExcento"].Value;
										this.txtIva.Value = (decimal)dr.Cells["SubtotalIva"].Value;
										this.txtDescIva0Total.Value = (decimal)dr.Cells["Descuento0"].Value;
										this.txtDescIvaTotal.Value = (decimal)dr.Cells["Descuento"].Value;
										this.txtIvaTotal.Value = (decimal)dr.Cells["Iva"].Value;
										this.txtTotal.Value = (decimal)dr.Cells["Total"].Value;

										dtFechaRecepcion = (DateTime)dr.Cells["FechaRecepcion"].Value;
										this.dtFecha.CalendarInfo.MinDate = (DateTime)dr.Cells["Fecha"].Value;

										if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
											this.txtIdTipoDePago.Value = 2;
										else
											this.txtIdTipoDePago.Value = 1;

										string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumidOrigen.Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo);

										iCont++;
										#endregion datos
									}
									else
									{
										#region datos2
										this.txtNumidOrigen1.Value = (int)dr.Cells["idCompra"].Value;
										string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumidOrigen1.Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo);

										this.txtIva0.Value = (decimal)this.txtIva0.Value + (decimal)dr.Cells["SubtotalExcento"].Value;
										this.txtIva.Value = (decimal)this.txtIva.Value + (decimal)dr.Cells["SubtotalIva"].Value;
										this.txtDescIva0Total.Value = (decimal)this.txtDescIva0Total.Value + (decimal)dr.Cells["Descuento0"].Value;
										this.txtDescIvaTotal.Value = (decimal)this.txtDescIvaTotal.Value + (decimal)dr.Cells["Descuento"].Value;
										this.txtIvaTotal.Value = (decimal)this.txtIvaTotal.Value + (decimal)dr.Cells["Iva"].Value;
										this.txtTotal.Value = (decimal)this.txtTotal.Value + (decimal)dr.Cells["Total"].Value;

										this.txtNotas.Text += " - " + dr.Cells["Numero"].Value.ToString();
										#endregion datos2
									}

									this.ultraGrid2.Rows.Band.AddNew();
									this.ultraGrid2.ActiveRow.Cells["idOrdenDeCompraCompra"].Value = 0;
									this.ultraGrid2.ActiveRow.Cells["idOrdenDeCompra"].Value = (int)dr.Cells["idCompra"].Value;
									this.ultraGrid2.ActiveRow.Cells["idCompra"].Value = (int)this.txtNumIdCompra.Value;

									if (sIdCompra.Length == 0)
										sIdCompra = dr.Cells["idCompra"].Value.ToString();
									else
										sIdCompra = sIdCompra + ", " + dr.Cells["idCompra"].Value.ToString();
								}
							}

							this.cmbFormaPago.Value = 6;
							this.txtNumEstado.Value = 7;
							this.dtFecha.Value = DateTime.Today;

							if ((int)this.txtIdGrupoCliente.Value != 3)
								this.txtNumEstado.Value = 7;


							//-------------------------------------------------------------
							// VALIDAR PRODUCTOS POR FACTURAR
							//-------------------------------------------------------------

							int estadoCabecera = 0;

							try
							{
								estadoCabecera = Convert.ToInt32(
									Funcion.iEjecucion(
									MenuLatinium.sconexionEmpresa,
									string.Format(
									"SELECT TOP 1 c.Estado FROM Compra c WHERE c.idCompra IN ({0})",
									sIdCompra
									),
									false
									)
									);
							}
							catch { }

							// SOLO valida en estos estados
							if (estadoCabecera == frmIngresoSeriales.EstadoPorFacturar ||
								estadoCabecera == frmIngresoSeriales.EstadoRecepcionBodega)
							{
								object oCantidad =
									Funcion.iEjecucion(
									MenuLatinium.sconexionEmpresa,
									string.Format(
									@"
            SELECT COUNT(*)
            FROM DetCompra dc
                INNER JOIN Compra c
                    ON c.idCompra = dc.idCompra
            WHERE dc.idCompra IN ({0})
            AND ISNULL(dc.EstadoFacturacion,0) = 1
            AND ISNULL(dc.CantidadPorFacturar,0) > 0
        ",
									sIdCompra
									),
									false
									);

								int iCantidad = 0;

								try
								{
									iCantidad = Convert.ToInt32(oCantidad);
								}
								catch { }

								if (iCantidad == 0)
								{
									MessageBox.Show(
										"La orden de compra no tiene productos marcados para facturar.",
										"Point Technology",
										MessageBoxButtons.OK,
										MessageBoxIcon.Stop
										);

									return;
								}
							}


							//-------------------------------------------------------------
							// CARGAR DETALLE
							//-------------------------------------------------------------

							this.ultraGrid1.DataSource =
								Funcion.dvProcedimiento(
								cdsSeteoF,
								string.Format(
								"Exec ConsultaDetalleOrdCompraACompra '{0}'",
								sIdCompra
								)
								);

							bNuevo = true; bEdicion = true;
							this.cmbBodega.Value = IBodega;
							this.dtFechadeEmicion.Enabled = true;

							if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
							{
								this.txtSerieFact.Enabled = true;
								this.txtAutFactura.Enabled = true;
								this.dtCaducidad.Enabled = true;
							}

							if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
							{
								this.cmbCreditoTributario.Value = 7;
								this.txtNumero.MaxLength = 20;
								this.txtIdTipoDePago.Value = 2;
							}
							else
							{
								this.cmbCreditoTributario.Value = 1;
								this.txtNumero.MaxLength = 9;
								this.txtIdTipoDePago.Value = 1;
							}

							this.txtNumero.Enabled = true;

							if ((int)this.txtIdGrupoCliente.Value != 3)
							{
								this.txtDiasCredito1.Enabled = false;
								this.txtDiasCredito2.Enabled = false;
								this.txtDiasCredito3.Enabled = false;
								this.txtDiasCredito4.Enabled = false;
								this.txtDiasCredito5.Enabled = false;
							}

							this.cmbTipoFormaPago.Enabled = true;

							if (miAcceso.CambiarFecha)
								this.dtFecha.Enabled = true;

							this.txtNotas.Enabled = true;
							this.btnNuevo.Enabled = false;
							this.btnConsultar.Enabled = false;
							this.btnGuardar.Enabled = true;

							if (miAcceso.BFTemporal)
								this.btnTemporal.Enabled = true;

							this.btnCancelar.Enabled = true;
							this.btnBuscarOrdenDeCompra.Enabled = false;
							//	this.btnSubir.Visible = true;

							this.ModificaGrilla(true);

							bNuevo = true; bEdicion = true;

							if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
								this.txtNumero.Focus();
							else
								this.txtAutFactura.Focus();

							this.cmdTotal();

						}
						catch (Exception Exc)
						{
							MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaTransacciones Busqueda = new frmBuscaTransacciones((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);						
				}
			}
		}

		private void cmbTipoFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtNumIdCompra.Value);

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				string sMensaje = "ANULAR";
				if (bAnulado) sMensaje = "ELIMINAR";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede '{1}' esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper(), sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}
			else if (bAnulado)
			{
				#region Eliminar
				if (miAcceso.Eliminar)
				{
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de eliminar este documento.", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						try
						{
							string sSQLBorrar = string.Format("Exec BorrarCompraVenta {0}, {1}, {2}", (int)this.txtNumIdCompra.Value, (int)this.txtNumIdAsiento.Value, (int)this.txtIdTipoFactura.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLBorrar, true);

							MessageBox.Show("Documento Eliminado del sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.btnCancelar_Click(sender, e);
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}						
					}
				}
				else MessageBox.Show("No tiene Acceso a Eliminar Facturas", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				#endregion Eliminar
			}
			else
			{
				#region Anular
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular Facturas De Compra", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else if (!miAcceso.BAnularPagosAutorizados && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos Where Estado = 1 And idTipo = 4 And idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
				{
					MessageBox.Show("No tiene acceso a Anular Facturas De Compra", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				}	
				else if (!miAcceso.BAnularPagosRealizados && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AutorizacionDePagos a Inner Join HistorialAutorizacionDePago h On h.idAutorizacion = a.idAutorizacionPago Where h.Estado = 1 And a.idTipo = 4 And a.idOrigen = {0}", (int)this.txtNumIdCompra.Value)) > 0)
				{
					MessageBox.Show("No tiene acceso a Anular Facturas De Compra", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				}
				else if (!miAcceso.BAnularLiquidado && (int)this.txtNumEstado.Value == 8)
				{
					MessageBox.Show("No tiene acceso a Anular Facturas De Compra", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);														
				}				
				else if (!Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF))			
				{
					if ((int)this.txtNumidOrigen.Value > 0 || (int)this.txtNumidOrigen1.Value > 0)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(ManejaSerial, 0) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
							{
								if (!Funcion.bEscalarSQL(cdsSeteoF, 
									string.Format("Declare @Disponible Bit If Exists(Select Disponible From DetalleSeriales Where idDetCompra = {0}) Select @Disponible = Disponible From DetalleSeriales Where idDetCompra = {0} Else Set @Disponible = 0 Select @Disponible", (int)dr.Cells["idDetCompra"].Value), true))	
								{
									MessageBox.Show(string.Format("Las seriales del producto '{0}' ya no estan disponibles en este documento", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.ultraGrid1.ActiveRow = dr;
									this.ultraGrid1.ActiveRow.Selected = true;
									return;
								}
							}
						}
					}

					if((int)this.cmbComprobante.Value == 3)
					{
						if(!miAcceso.AnularDocSRI)
						{
							if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFE_EstadoComprobante From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true) == 2)
							{
								MessageBox.Show("No puede Anular la liquidación de compra  esta autorizada por el S.R.I.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}
					}

					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular esta Factura", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{
						using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
						{
							if (DialogResult.OK == miNota.ShowDialog())
							{
								try
								{
									string sSQLAnul = string.Format("Exec AnulacionCompraVenta {0}, 4, '{1}'", (int)this.txtNumIdCompra.Value, miNota.txtNotas.Text.ToString());
									Funcion.sEscalarSQL(cdsSeteoF, sSQLAnul);

									this.Consultar((int)this.txtNumIdCompra.Value);

									bAnulado = true;

									this.lblEstado.Text = "ANULADO";

									if (bAnulado) this.btnAnular.Text = "Eliminar";

								}
								catch(Exception Exc)
								{
									MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
							}
						}
					}
				}
				#endregion Anular
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{			
			DateTime dtFechaAsi = (DateTime)dtFecha.Value;

			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (bAnulado) 
			{
				MessageBox.Show("No puede crear un Asiento de una Factura Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				MessageBox.Show(string.Format("No Puede Crear Asiento de esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFechaAsi.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}						
			else
			{					
				string stExisteAsiento = string.Format("Exec AsientoMaestroExiste {0}", (int)this.txtNumIdCompra.Value);
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
			
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);					
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
							if (iOrigen < 2)
							{
								string stGeneraCompra = string.Format("Exec CreaAsiento {0}", (int)this.txtNumIdCompra.Value);
								Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);
							}
							else if (iOrigen == 2)
							{
								string sSQLAsiento = string.Format("Exec AsientoCompraCTC {0}", (int)this.txtNumIdCompra.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
							}

							this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true);

							MessageBox.Show("Asiento Generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			

							Cursor = Cursors.Default;
							#endregion Crea Asiento
						}
					}
					else if (DialogResult.Yes == MessageBox.Show("¿Desea Crear Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{
						#region Crea Asiento
						if (iOrigen < 2)
						{
							string stGeneraCompra = string.Format("Exec CreaAsiento {0}", (int)this.txtNumIdCompra.Value);
							Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);
						}
						else if (iOrigen == 2)
						{
							string sSQLAsiento = string.Format("Exec AsientoCompraCTC {0}", (int)this.txtNumIdCompra.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
						}

						this.txtNumIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idAsiento From Compra Where idCompra = {0}", (int)this.txtNumIdCompra.Value), true);

						MessageBox.Show("Asiento Generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			

						Cursor = Cursors.Default;
						#endregion Crea Asiento
					}
				}								
			}
		}

		private void Imprimir(int iTipo)
		{
			Funcion.TiempoInicia();			
			DateTime dtIni = DateTime.Now;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + this.Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if (!miAcceso.ReImprimir && bImpreso)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Compra.rpt";
			
			if (MenuLatinium.stDirFacturacion == "CROMATIKA" || 
				MenuLatinium.stDirFacturacion == "LOTESERVICIO" || 
				MenuLatinium.stDirFacturacion == "GRETASANTORINI" ||
				MenuLatinium.stDirFacturacion == "PROSERVICIOS" ||
				MenuLatinium.stDirFacturacion == "BECAJAJUMA" ||
				MenuLatinium.stDirFacturacion == "SERVIFLASH" ||
				MenuLatinium.stDirFacturacion == "CUEROFINO" ||
				MenuLatinium.stDirFacturacion == "FINTECHLIFE" ||
				MenuLatinium.stDirFacturacion == "KARPATOS" ||
				MenuLatinium.stDirFacturacion == "ESPARTA") stReporte = "CompraCromatika.rpt";

			/*string.Format("Exec CompraReporteNombre {0}, 0", (int)this.txtNumIdCompra.Value);							
				stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);*/
			string stFiltro = "{Compra.idCompra} = " + this.txtNumIdCompra.Value.ToString();
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				if (iTipo == 0) miReporte.ShowDialog(); // Muestra en pantalla

				if (iTipo >= 1) // Imprime una o mas copias
				{
					string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
					string stImpresora = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraReporteImpresora '" + stReporte + "'");
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
					miReporte.ImprimeRep(1);
					if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
				}

				if (iTipo == -1) miReporte.ImprimeRep(2);// Archivo PDF
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = string.Format("Update Compra Set Impreso = 1 Where idCompra = {0}", (int)this.txtNumIdCompra.Value);
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			bImpreso = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			Imprimir(0);
			int iCont = 0;
				
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
				int iSerial = 0;
				iSerial = Funcion.iEscalarSQL(cdsSeteoF, string.Format("IF (( Select idTipoGrupo From Articulo Where idArticulo = {0}) = 1)BEGIN sELECT CONVERT (INT,1) END ELSE BEGIN sELECT CONVERT (INT,0)END", (int)dr.Cells["idArticulo"].Value));	
				if (iSerial ==1) iCont = iCont +1;								
			}
			if (iCont > 0)
			{
				string stFiltro = "{Compra.idCompra} = " + this.txtNumIdCompra.Value.ToString();			

				Reporte miReporte = new Reporte("HojaDeSeriales.rpt", stFiltro);
				miReporte.ShowDialog();
			}
		}

		private void btnImprimir_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			#region Numero de Copias de CompraNumero
			int iCopias = Funcion.iEscalarSQL(cdsSeteoF, "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = 4");
			if (iCopias < 1) iCopias = 1;
			#endregion Numero de Copias de CompraNumero

			Imprimir(iCopias);
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
							frmAsientos Asientos = new frmAsientos((int) this.txtNumIdAsiento.Value);
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

		private void btnRetencion_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtNumIdCompra.Value);

			if (!miAcceso.Retencion) 
			{
				MessageBox.Show("No tiene Acceso a Ver Retenciones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}			
			else if (bAnulado) 
			{
				MessageBox.Show("La Factura esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (iOrigen == 2) 
			{
				MessageBox.Show("La Factura no Genera Retencion La Forma de Pago es Tarjeta de Crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}			
			else if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idRetencion From Retenciones Where ClienteProveedor = 1 And idCompra = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtNumIdCompra.Value), true)) 
			{
				MessageBox.Show("No existe Retención de esta Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				frmRetenciones Retenciones = new frmRetenciones(this.ultraGrid1, (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, this.txtNumero.Text.ToString(), (DateTime)this.dtFecha.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, false, (int)this.cmbComprobante.Value, bLiquidacion);
       				
				Retenciones.ShowDialog();
			}
		}

		private void cmbCreditoTributario_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{ 				
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}
			
			if (e.Cell.Column.ToString() == "Articulo")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "Precio")
			{ 				
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}
			
			if (e.Cell.Column.ToString() == "Cantidad")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "Impuesto")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}
			
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "Valores")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "DescNC")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "idProyecto")
			{ 
				if ((int)e.Cell.Row.Cells["idOrigen"].Value > 0)
				{
					e.Cancel = true;
				}
			}
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDiasCredito1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDiasCredito2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDiasCredito3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkConsignacion_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerieFact_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{				
				if (miAcceso.CambiarFecha) this.dtFecha.Focus();
				else this.txtDiasCredito1.Focus();				
			}
		}

		private void txtAutFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtSerieFact.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)  
			{
				this.dtFechadeEmicion.Focus();
			}
		}

		private void dtCaducidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.txtDiasCredito1.Focus();
		}

		private void txtDiasCredito1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiasCredito2.Focus();
		}

		private void txtDiasCredito2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiasCredito3.Focus();
		}

		private void txtDiasCredito3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiasCredito4.Focus();
		}

		private void cmbTipoFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{	
				if (this.cmbImportacionNombre.Enabled)this.cmbImportacionNombre.Focus();
				else
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
				}
			}
		}

		private void cmbImportacionNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.cmbCreditoTributario.Enabled) this.cmbCreditoTributario.Focus();
				else
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
			}
		}

		private void cmbCreditoTributario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.ultraGrid1.ActiveRow.Cells["idOrigen"].Value > 0 && 
				(this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "PRECIO" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "IMPUESTO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "DESCUENTOPORC")) e.Handled = true;			
		}

		private void frmCompra_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void cmbImportacionNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 && (int)this.txtIdGrupoCliente.Value == 3)
				{
					if (this.cmbImportacionNombre.ActiveRow != null)
					{
						MessageBox.Show("Seleccione una Importación, PROVEEDOR INTERNACIONAL O DE ASIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbImportacionNombre.Focus();
						e.Cancel = true;
					}
				}
			}
		}

		private void frmCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) this.btnConsultar_Click(sender, e);
				if ((int)this.txtNumIdCompra.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) this.btnImprimir_Click(sender, e);
					if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) this.btnAnular_Click(sender, e);
					if (e.KeyCode == Keys.X) if (miAcceso.Anular && bAnulado) this.btnAnular_Click(sender, e);
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0603");
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
			}
		}

		private void cmbImportacionNombre_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbImportacionNombre.ActiveRow != null) this.ultraGrid1.DisplayLayout.Bands[0].Columns["idImportacion"].Hidden = false;
			else this.ultraGrid1.DisplayLayout.Bands[0].Columns["idImportacion"].Hidden = true;
		}

		private void txtNumContadoCredito_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				ActualizaIVA();
			}
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtCaducidad_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void txtDiasCredito4_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito4.SelectAll();
		}

		private void txtDiasCredito4_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito4.SelectAll();
		}

		private void txtDiasCredito4_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoFormaPago.Focus();
		}

		private void btnProforma_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaProformasAutorizadas Busqueda = new frmConsultaProformasAutorizadas())
			{ 
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtNumidOrigen.Value = (int)Busqueda.uGridProformas.ActiveRow.Cells["idProforma"].Value;					
					this.txtNumContadoCredito.Value = 2;
					this.txtNumIdCliente.Value = (int)Busqueda.uGridProformas.ActiveRow.Cells["idCliente"].Value;
					this.txtIdTipoRuc.Value = (int)Busqueda.uGridProformas.ActiveRow.Cells["idTipoRuc"].Value;
					this.txtIdGrupoCliente.Value = (int)Busqueda.uGridProformas.ActiveRow.Cells["idGrupoCliente"].Value;
					this.cmbComprobante.Value = 1;
					this.txtRuc.Text = Busqueda.uGridProformas.ActiveRow.Cells["Ruc"].Value.ToString();
					this.txtNombre.Text = Busqueda.uGridProformas.ActiveRow.Cells["Nombre"].Value.ToString();
					dValorProforma = Convert.ToDecimal(Busqueda.uGridProformas.ActiveRow.Cells["Total"].Value);

					if ((int)this.txtIdGrupoCliente.Value == 2) this.txtIdTipoDePago.Value = 2;
					else if ((int)this.txtIdGrupoCliente.Value == 3)  this.txtIdTipoDePago.Value = 2;
					else this.txtIdTipoDePago.Value = 1;

					this.cmbFormaPago.Value = 6;
					this.txtNumEstado.Value = 7;
					this.dtFecha.CalendarInfo.MinDate = Convert.ToDateTime(Busqueda.uGridProformas.ActiveRow.Cells["Fecha"].Value);
					this.dtFecha.Value = DateTime.Today;

					if ((int)this.txtIdGrupoCliente.Value != 3) this.txtNumEstado.Value = 7;
										
					bNuevo = true;
					bEdicion = true;

					this.cmbBodega.Value = 1;

					if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3)
					{
						this.txtSerieFact.Enabled = true;
						this.txtAutFactura.Enabled = true;
						this.dtCaducidad.Enabled = true;								
					}

					if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
					{
						this.cmbCreditoTributario.Value = 7;
						this.txtNumero.MaxLength = 20;
						this.txtIdTipoDePago.Value = 2;
					}
					else 
					{
						this.cmbCreditoTributario.Value = 1;
						this.txtNumero.MaxLength = 9;
						this.txtIdTipoDePago.Value = 1;
					}

					this.txtNumero.Enabled = true;
						
					if ((int)this.txtIdGrupoCliente.Value != 3)
					{
						this.txtDiasCredito1.Enabled = false;
						this.txtDiasCredito2.Enabled = false;
						this.txtDiasCredito3.Enabled = false;
						this.txtDiasCredito4.Enabled = false;
						this.txtDiasCredito5.Enabled = false;
					}

					this.cmbTipoFormaPago.Enabled = true;
						
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
											
					this.txtNotas.Enabled = true;
						
					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnGuardar.Enabled = true;
					if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
					this.btnCancelar.Enabled = true;
					this.btnBuscarOrdenDeCompra.Enabled = false;
					this.btnProforma.Enabled = false;
						
					this.ModificaGrilla(true);

					bNuevo = true;
					bEdicion = true;
					
					if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.txtNumero.Focus(); else this.txtAutFactura.Focus();

					bProforma = true;					 
				}
			}
		}


		
	

		public string NuevoGoogle(string txtDirectorioarchivo )
		{
			if ( !File.Exists(txtDirectorioarchivo))
			{
				MessageBox.Show("El archivo seleccionado no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}

			try
			{
				// Configuraci?n de la ruta destino
				string carpetaDestino = @"\\192.168.1.10\Proyectos\Imagenes\" 
					+ MenuLatinium.stDirFacturacion 
					+ @"\Compras" 
					+ @"\" 
					+ (string)this.txtRuc.Text;

     
				// Crear carpeta si no existe
				if (!Directory.Exists(carpetaDestino))
				{
					Directory.CreateDirectory(carpetaDestino);
				}

				// Obtener info del archivo original
				FileInfo fi = new FileInfo(txtDirectorioarchivo);
				string extension = fi.Extension;

				// Generar nombre ?nico con timestamp
				string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
				string nuevoNombreArchivo = "archivo_" + timestamp + extension;
				sCarpeta = nuevoNombreArchivo;



				string rutaDestinoFinal = Path.Combine(carpetaDestino, nuevoNombreArchivo);

				// Si ya existe un archivo con ese nombre, lo elimina (altamente improbable por timestamp)
				if (File.Exists(rutaDestinoFinal))
				{
					File.Delete(rutaDestinoFinal);
				}

				// Conexi?n a red (forzar acceso)
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p");

				// Mover archivo
				File.Copy(txtDirectorioarchivo, rutaDestinoFinal);

				// Confirmaci?n opcional
				return rutaDestinoFinal;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
		}

		private void btnSubir_Click(object sender, System.EventArgs e)
		{
		

			/* Cargar imagen */
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.InitialDirectory = @"C:\";
			ofd.Filter = "Archivos PDF (*.pdf)|*.pdf";

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				string archivoSeleccionado = ofd.FileName;
				

				// Llama al m?todo y obtiene la nueva ruta
				string nuevaRuta = NuevoGoogle(archivoSeleccionado);

				if ( nuevaRuta == null || nuevaRuta == "")
				{
					
					MessageBox.Show("Hubo un problema al subir la imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					sUrlNube = nuevaRuta;
				}
			}
			else
			{
				MessageBox.Show("No se seleccion? una imagen.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}


	
		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (sUrlNube.Length > 10 )
			{
				//Process.Start(sUrlNube);
				//	System.Diagnostics.Process.Start("Google", sUrlNube);

				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}
		}

		private void label13_Click(object sender, System.EventArgs e)
		{
		
		}

		private void dtFechadeEmicion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (MenuLatinium.stDirFacturacion == "POINT")
			{
				if (bEdicion)
				{
					if (this.dtFechadeEmicion.Value == DBNull.Value)
					{
						MessageBox.Show("Ingrese la Fecha de Emisión.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.dtFechadeEmicion.Focus();
						e.Cancel = true;				
					}
					else if (Funcion.ValidaPeriodo((DateTime)this.dtFechadeEmicion.Value, cdsSeteoF, "Compra"))
					{
						DateTime dtFechadeEmicion = (DateTime)this.dtFechadeEmicion.Value;
						MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFechadeEmicion.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.dtFecha.Focus();
						e.Cancel = true;					
					}
				}		
			}
			//else {FuncionesProcedimientos.RetornaFechaServidor(this.dtFechadeEmicion, cdsSeteoF, false);}
		}

		private void dtFechadeEmicion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFechadeEmicion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)  
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.txtDiasCredito1.Focus();
				else this.cmbCreditoTributario.Focus();
			}
		}

		private void btnDistribucion_Click(object sender, System.EventArgs e)
		{
			#region Validacion Maestro
			//string Periodo  = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			if (this.txtRuc.Text.Length == 0 )
			{
				this.txtAutFactura.Focus();
				return;				
			}

			if(!bLiquidacion)
			{
				if (this.txtSerieFact.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;				
				}
				if (this.txtNumero.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el número del Documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;				
				}

				if (this.txtAutFactura.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					return;				
				}
			}
			if (this.dtCaducidad.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
			{
				MessageBox.Show("Seleccione la Fecha de Caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();
				return;				
			}
			#endregion Validacion Maestro

			this.mnuDistribucion.Show(this.btnDistribucion, new Point(1, 20));
		}

		private void txtNumEstado_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtFechaPeriodo_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void mnuAgencias_Click(object sender, System.EventArgs e)
		{
			Distribucion(1);
		}

		private void mnuAdministrativos_Click(object sender, System.EventArgs e)
		{
			Distribucion(2);
		}

		private void mnuGeneral_Click(object sender, System.EventArgs e)
		{
			Distribucion(3);
		}

		public void  Distribucion (int sTipo)
		{
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
			#region  distribucion 
			if ((this.txtAutFactura.Text.Length >0 && this.txtRuc.Text.Length > 0)|| bLiquidacion)
			{
				using (frmCompraDistribucion ComprasDistribucion = new frmCompraDistribucion())
				{
					if (DialogResult.OK == ComprasDistribucion.ShowDialog())
					{
						if (ComprasDistribucion.uGridDetalle.Rows.Count > 0)
						{	
							#region comentado
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ComprasDistribucion.uGridDetalle.Rows.All)
							{
								int iCuenta =(int)dr.Cells["idCuenta"].Value ;	
								decimal iValor  = (decimal)dr.Cells["Valor"].Value;
								try
								{	
									int IBodega = 1;
									this.cmbFormaPago.Value = 6;
									this.txtNumEstado.Value = 7;
									this.dtFecha.Value = DateTime.Today;

									if ((int)this.txtIdGrupoCliente.Value != 3) this.txtNumEstado.Value = 7;
													
									string sSQLInf = string.Format("Exec GeneraCompraDistribucion {0},{1},'{2}',{3}", iValor,iCuenta,dtFechaFac.ToString("yyyyMMdd"),sTipo);
									SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
									while (drInf.Read())
									{	
										bActualiza=false;
										this.ultraGrid1.Rows.Band.AddNew();
										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drInf.GetString(5);
										this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = drInf.GetString(6);
										this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value = drInf.GetInt32(7);
										this.ultraGrid1.ActiveRow.Cells["Precio"].Value = drInf.GetDecimal(8);
										this.ultraGrid1.ActiveRow.Cells["Impuesto"].Value = drInf.GetDecimal(9);
										this.ultraGrid1.ActiveRow.Cells["idProyecto"].Value = drInf.GetInt32(13);
										this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = drInf.GetInt32(3);
										bActualiza=true;
									}
									drInf.Close();	
									cmdTotal();

									bNuevo = true;
									bEdicion = true;

									#region Controles
									this.cmbBodega.Value = IBodega;
									this.dtFechadeEmicion.Enabled = true;

									this.btnNuevo.Enabled = false;
									this.btnConsultar.Enabled = false;
									this.btnGuardar.Enabled = true;
									if (miAcceso.BFTemporal) this.btnTemporal.Enabled = true;
									this.btnCancelar.Enabled = true;
									this.btnBuscarOrdenDeCompra.Enabled = false;
									this.btnDistribucion.Enabled = false;
									//this.btnSubir.Visible = true;
									#endregion Controles
						
									bNuevo = true;
									bEdicion = true;
					
									if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.txtNumero.Focus(); else this.txtAutFactura.Focus();
									this.ultraGrid1.Rows.Band.AddNew();
								}
								catch(Exception Exc)
								{
									if (this.ultraGrid1.ActiveRow.Cells["Codigo"].Value == System.DBNull.Value)
									{
										MessageBox.Show("Cuenta no asignada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
										return;
									}
								
									MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

								}
							}
							#endregion comentado
								
							bEliminaAlValidar = true;
							Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
							bEliminaAlValidar = false;

						}
					}
				}
			}
			#endregion  distribucion 
		}

		private void txtRuc_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnTemporal_Click(object sender, System.EventArgs e)
		{
			#region Validacion Encabezado
			#region Busca Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, (bool)this.chkConsignacion.Checked);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (!drC.HasRows)
			{
				drC.Close();		
				MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			drC.Close();		
			#endregion Busca Proveedor
			#region E-Mail
			string sEMail = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Email From Cliente Where idCliente = {0} And Proveedor = 1", (int)this.txtNumIdCliente.Value));
			if (!Validacion.bEMail(sEMail.Trim())) return;
			#endregion E-Mail
			#region Serie
			if (!bLiquidacion)
			{
				if (this.txtSerieFact.Text.ToString().Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}

				if (this.txtSerieFact.Text.Length < 6 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerieFact.Focus();
					return;
				}
			}
			#endregion Serie
			#region Numero
			if (!bLiquidacion)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
				if (this.txtNumero.Text.ToString().Length <9)
				{
				
					MessageBox.Show("La longitud del número del documento debe ser de 9 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeracionCompra '{0}', '{1}', '{2}', {3}, {4}, 4, {5}", 
					this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, (int)this.cmbComprobante.Value), true);
				if (bRes)
				{
					MessageBox.Show("El Número de Documento ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();				
					return;
				}
			}
			#endregion Numero
			#region Autorizacion
			if (!bLiquidacion)
			{
				if (this.txtAutFactura.Text.Length == 0 && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					return;
				}
				if ((this.txtAutFactura.Text.Length != 10 && this.txtAutFactura.Text.Length != 37 && this.txtAutFactura.Text.Length != 49) && (int)this.txtIdTipoRuc.Value != 3)
				{
					MessageBox.Show("Longitud de autorización debe tener 10, 37 o 49 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtAutFactura.Focus();
					return;
				}
			}
			#endregion Autorizacion
			#region Fecha
			if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Fecha
			#region Caducidad
			if (this.dtCaducidad.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
			{
				MessageBox.Show("Seleccione la Fecha de Caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();
				return;
			}
			if (MenuLatinium.stDirFacturacion == "DENTALIS")
			{
				if (this.dtFechaPeriodo.Value == System.DBNull.Value )
				{
					MessageBox.Show("Seleccione la Fecha del Periodo de la Compra.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFechaPeriodo.Focus();
					return;
				}
			}

			if ((int)this.txtIdTipoRuc.Value != 3)
			{
				DateTime dtValida = (DateTime) this.dtFecha.Value;
				dtValida = dtValida.AddDays(1);

				if ((DateTime)this.dtCaducidad.Value < dtValida)
				{
					MessageBox.Show("Fecha de caducidad no valida.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtCaducidad.Focus();
					return;
				}
			}
			#endregion Caducidad
			#region FechadeEmicion
			if (this.dtFechadeEmicion.Value == System.DBNull.Value && (int)this.txtIdTipoRuc.Value != 3)
			{
				MessageBox.Show("Seleccione la Fecha de Emisión.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechadeEmicion.Focus();
				return;
			}
			#endregion FechadeEmicion
			#region Dias De Credito
			//			if ((int)this.txtIdGrupoCliente.Value != 2 || (int)this.txtIdGrupoCliente.Value != 3)
			//			{
			//				if ((int)this.txtDiasCredito1.Value == 0 && (int)this.txtDiasCredito2.Value == 0 && (int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value == 0)
			//				{
			//					MessageBox.Show("Ingrese los dias de credito para el pago del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito1.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito1.Value == 0 && ((int)this.txtDiasCredito2.Value > 0 || (int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
			//				{
			//					MessageBox.Show("Dias Credito 1 no puede estar en Cero si Dias Credito 2, 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito1.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito2.Value == 0 && ((int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
			//				{
			//					MessageBox.Show("Dias Credito 2 no puede estar en Cero si Dias Credito 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito2.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value > 0)
			//				{
			//					MessageBox.Show("Dias Credito 3 no puede estar en Cero si Dias Credito 4 es mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito3.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito2.Value > 0 && (int)this.txtDiasCredito1.Value > (int)this.txtDiasCredito2.Value)
			//				{
			//					MessageBox.Show("Dias Credito 1 no puede ser mayor a Dias de Credito 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito2.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito3.Value > 0 && (int)this.txtDiasCredito2.Value > (int)this.txtDiasCredito3.Value)
			//				{
			//					MessageBox.Show("Dias Credito 2 no puede ser mayor a Dias de Credito 3", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito3.Focus();
			//					return;
			//				}
			//				if ((int)this.txtDiasCredito4.Value > 0 && (int)this.txtDiasCredito3.Value > (int)this.txtDiasCredito4.Value)
			//				{
			//					MessageBox.Show("Dias Credito 3 no puede ser mayor a Dias de Credito 4", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDiasCredito4.Focus();
			//					return;
			//				}
			//			}
			#endregion Dias De Credito
			#region Forma de Pago SRI
			if ((decimal)this.txtTotal.Value > 1000.00m)
			{
				if (this.cmbTipoFormaPago.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una forma de pago.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipoFormaPago.Focus();
					return;
				}
			}
			#endregion Forma de Pago SRI
			#region Importacion
			if ((int)this.txtIdGrupoCliente.Value == 2 && (int)this.txtIdGrupoCliente.Value == 3)
			{
				if (this.cmbImportacionNombre.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una importación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbImportacionNombre.Focus();
					return;
				}
			}
			#endregion Importacion
			#region CreditoTributario
			if (this.cmbCreditoTributario.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Crédito Tributario para este documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCreditoTributario.Focus();
				return;
			}
			#endregion CreditoTributario
			#region Transaccion en Cero
			if ((decimal) this.txtTotal.Value == 0) 
			{
				MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Transaccion en Cero			
			#endregion Validacion Encabezado

			#region Validacion Detalle
			#region Importacion
			if ((int)this.txtIdGrupoCliente.Value == 2 && (int)this.txtIdGrupoCliente.Value == 3)
			{
				if (this.cmbImportacionNombre.ActiveRow != null)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if (dr.Cells["idImportacion"].Value == DBNull.Value)
						{
							MessageBox.Show(string.Format("Seleccione el tipo de importación para : '{0}'", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
				}
			}
			#endregion Importacion
			#region Valida Cantidad - Precio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int IdArt = (int) dr.Cells["idArticulo"].Value;
					decimal dPrecio = (decimal) dr.Cells["Precio"].Value;

					if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value) 
					{
						MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
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
			#region Valida Centro de Costo
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idProyecto"].Value == DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione el Centro de Costo : '{0}'", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Valida Centro de Costo		
			#region Google Nuevo
			if (MenuLatinium.stIdDB != 1 && MenuLatinium.stIdDB != 33&& MenuLatinium.stIdDB != 3 )
			{
				if  ( (Decimal) this.txtTotal.Value >= 1000)
				{
					if(!bNuevo)
					{
						if (sUrlNube.Length < 10 )
						{
							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
					else
					{
						if (txtDirectorioarchivo.Length < 10 )
						{
							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
			}
			else
			{
				//if  ( (Decimal) this.txtTotal.Value > 5000)
				if ((int)this.txtIdGrupoCliente.Value == 1 ||((int)this.txtIdGrupoCliente.Value == 5 && (Decimal)this.txtTotal.Value >= 5000))				
				{
					if(!bNuevo)
					{
						if (sUrlNube.Length < 10 )
						{
							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
					else
					{
						if (txtDirectorioarchivo.Length < 10 )
						{
							MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
			}
			#endregion Google Nuevo

			#endregion Validacion Detalle

			int idProyecto = 0;
			if (this.cmbProyecto.ActiveRow != null) idProyecto = (int)this.cmbProyecto.Value;
			cmdTotal();

			bool bNumeracionAutomatica = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Automatico From CompraNumero Where idTipoFactura = 85"), true); 

			#region Valida Retencion
			bool bRetencion = false;
			if (this.txtRuc.Text.ToString().Trim() == Funcion.sEscalarSQL(cdsSeteo, "Select LTRIM(RTRIM(Ruc)) From Seteo")) bRetencion = false;
			else if (Funcion.iEscalarSQL(cdsSeteoF, "Select idAgenteRetencion From Seteo") == 5) bRetencion = false;
			else if ((int)this.cmbComprobante.Value == 2 || (int)this.cmbComprobante.Value == 12) bRetencion = false;
			else if ((int)this.cmbTipoFormaPago.Value == 19) bRetencion = false;
			else
			{
				int iArtRenta = 0;
				int iArtIVA = 0;

				bool bRetieneIVA = false;
				if ((decimal)this.txtIvaTotal.Value > 0) bRetieneIVA = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select IsNull(RetieneIVA, 0) From Cliente Where idCliente = {0}", (int)this.txtNumIdCliente.Value), true);
				
				if ((int)this.cmbComprobante.Value == 1 || (int)this.cmbComprobante.Value == 3)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						int ArtRenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select IsNull(idArticuloSRI, 0) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
						if (ArtRenta != 0 && ArtRenta != 61) iArtRenta ++;
					}
				}

				if (bRetieneIVA || iArtRenta > 0) bRetencion = true;
				else bRetencion = false;
			}
			#endregion Valida Retencion

			#region Variables
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			DateTime dtFechaPago = (DateTime)this.dtFecha.Value;
			DateTime dtFechadeEmicion = (DateTime)this.dtFechadeEmicion.Value;
			DateTime dtCaducidad = DateTime.Today;
			DateTime dtFechaPEriodo  = DateTime.Parse("01/01/2000");
			bool bFecCad = false;
			int IdTipoPago = 0;
			int IdImportacionNombre = 0;
			int IdProyectoDetalle = 0;
			int IdImportacionTipo = 0;
			int idOrigen = 0;
			bool bDiasCompra = false;

			if (this.cmbTipoFormaPago.Value != System.DBNull.Value) IdTipoPago = (int)this.cmbTipoFormaPago.Value;
			if (this.cmbImportacionNombre.ActiveRow != null) IdImportacionNombre = (int)this.cmbImportacionNombre.Value;

			if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
			{
				dtCaducidad = (DateTime)this.dtCaducidad.Value;
				bFecCad = true;
			}
			if (this.dtFechaPeriodo.Value != System.DBNull.Value)dtFechaPEriodo=(DateTime)this.dtFechaPeriodo.Value;

			#endregion Variables

			#region Graba Registro
				
			if (bLiquidacion || (bRetencion && (int)this.cmbFormaPago.Value == 6))
			{
				#region Registro Con Retenciones
				using (frmRetenciones Retenciones = new frmRetenciones(this.ultraGrid1, (int)this.txtNumIdCompra.Value, (int)this.txtNumIdCliente.Value, this.txtNumero.Text.ToString(), (DateTime)this.dtFecha.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, true, (int)this.cmbComprobante.Value, bLiquidacion))
				{
					if (DialogResult.OK == Retenciones.ShowDialog())
					{
						using (frmDiasCompra DiasCompra = new frmDiasCompra((int)this.txtNumIdCompra.Value,1))
						{
							if (DialogResult.OK == DiasCompra.ShowDialog())
							{
								if (DiasCompra.uGridDetalle.Rows.Count > 0)
								{	
									#region Proceso con retencion

									#region Valida Pagos Realizados
									decimal dPagado = 0;
									decimal dAutorizado = 0;
									decimal dPA = 0;
									string sSQLValPagos = string.Format("Exec ValidaPagosCompras {0}", (int)this.txtNumIdCompra.Value);
									SqlDataReader drVP = Funcion.rEscalarSQL(cdsSeteoF, sSQLValPagos, true);
									drVP.Read();
									if (drVP.HasRows)
									{
										dPagado = drVP.GetDecimal(0);
										dAutorizado = drVP.GetDecimal(1);
									}
									drVP.Close();

									dPA = dPagado + dAutorizado;
									#endregion Valida Pagos Realizados

									if (Math.Round((decimal)this.txtTotal.Value, 2) == 0)// < Math.Round((dPA + (decimal)Retenciones.txtRetenido.Value), 2))
									{
										MessageBox.Show(string.Format("El Valor Pagado {0} más el Valor Autorizado {1} más el Valor Retenido {2} es Mayor al Valor Total de la Factura {3}", dPagado, dAutorizado, (decimal)Retenciones.txtRetenido.Value, (decimal)this.txtTotal.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									}
									else
									{
										#region Caso Contrario
										using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

												#region Proceso


												#region Graba Registro
												#region Graba Maestro
												int idProyectoZOOM = 0;
												this.txtNumEstado.Value = 6;
												if (this.cmbProyectoZOOM.ActiveRow != null) idProyectoZOOM = (int)this.cmbProyectoZOOM.Value;

												//												string sSQLMaestro = string.Format("Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}',{39},'{40}'",
												//													(int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtNumidOrigen1.Value,
												//													(int)this.txtNumIdAsiento.Value, (int)this.txtIdSucursal.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbComprobante.Value, 
												//													idProyectoZOOM, (int)this.txtNumContadoCredito.Value, (int)this.txtNumEstado.Value, (int)this.txtNumIdCliente.Value, 
												//													(int)this.cmbBodega.Value, idProyecto, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(),
												//													this.txtAutFactura.Text.ToString(), dtFecha.ToString("yyyyMMdd"), dtCaducidad.ToString("yyyyMMdd"),
												//													(int)this.cmbFormaPago.Value, (int)this.txtDiasCredito1.Value, (int)this.txtDiasCredito2.Value, (int)this.txtDiasCredito3.Value, IdTipoPago,
												//													IdImportacionNombre, this.txtComprobante.Text.ToString(), (bool)this.chkConsignacion.Checked, this.txtNotas.Text.ToString(),
												//													(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
												//													bFecCad, (int)this.txtIdTipoDePago.Value, (int)this.cmbCreditoTributario.Value, dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), (int)this.txtDiasCredito4.Value ,dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),(int)this.txtDiasCredito5.Value,dtFechadeEmicion.ToString("yyyyMMdd HH:mm"));			
												
												string sSQLMaestro = string.Format(
													"Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39}, '{40}'",

													(int)this.txtNumIdCompra.Value,
													(int)this.txtNumidOrigen.Value,
													(int)this.txtNumidOrigen1.Value,
													(int)this.txtNumIdAsiento.Value,
													(int)this.txtIdSucursal.Value,
													(int)this.txtIdTipoFactura.Value,
													(int)this.cmbComprobante.Value,
													idProyectoZOOM,
													(int)this.txtNumContadoCredito.Value,
													(int)this.txtNumEstado.Value,
													(int)this.txtNumIdCliente.Value,
													(int)this.cmbBodega.Value,
													idProyecto,

													LimpiarValorSQL(this.txtSerieFact.Text),
													LimpiarValorSQL(this.txtNumero.Text),
													LimpiarValorSQL(this.txtAutFactura.Text),

													dtFecha.ToString("yyyyMMdd"),
													dtCaducidad.ToString("yyyyMMdd"),

													(int)this.cmbFormaPago.Value,
													(int)this.txtDiasCredito1.Value,
													(int)this.txtDiasCredito2.Value,
													(int)this.txtDiasCredito3.Value,
													IdTipoPago,
													IdImportacionNombre,

													LimpiarValorSQL(this.txtComprobante.Text),

													(bool)this.chkConsignacion.Checked,
													LimpiarValorSQL(this.txtNotas.Text),

													(decimal)this.txtIva0.Value,
													(decimal)this.txtDescIva0Total.Value,
													(decimal)this.txtIva.Value,
													(decimal)this.txtDescIvaTotal.Value,
													(decimal)this.txtIvaTotal.Value,
													(decimal)this.txtTotal.Value,

													bFecCad,
													(int)this.txtIdTipoDePago.Value,
													(int)this.cmbCreditoTributario.Value,

													dtFechaRecepcion.ToString("yyyyMMdd HH:mm"),
													(int)this.txtDiasCredito4.Value,
													dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),
													(int)this.txtDiasCredito5.Value,
													dtFechadeEmicion.ToString("yyyyMMdd HH:mm")
													);
												
												oCmdActualiza.CommandText = sSQLMaestro;
												this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();//Funcion.iEscalarSQL(cdsSeteoF, sSQLMaestro, true);
												#endregion Graba Maestro

												#region Graba Detalle
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
												{	
													if (dr.Cells["idProyecto"].Value != System.DBNull.Value) IdProyectoDetalle = (int)dr.Cells["idProyecto"].Value; else IdProyectoDetalle = 0;
													if (dr.Cells["idImportacion"].Value != System.DBNull.Value) IdImportacionTipo = (int)dr.Cells["idImportacion"].Value; else IdImportacionTipo = 0;
													if (dr.Cells["idOrigen"].Value != System.DBNull.Value) idOrigen = (int)dr.Cells["idOrigen"].Value; else idOrigen = 0;
				
													string sSQLDetalle = string.Format("Exec GrabaDetalleCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}', '{23}'", 
														(int)dr.Cells["idDetCompra"].Value, idOrigen, (int)this.txtNumIdCompra.Value, 
														(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, Math.Round((decimal)dr.Cells["DescuentoArt"].Value, 2), 
														(decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value, (decimal)dr.Cells["DescNC"].Value,
														IdProyectoDetalle, IdImportacionTipo, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, 
														(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, (decimal)dr.Cells["Valores"].Value,
														(int)dr.Cells["idTipoGasto"].Value, dr.Cells["Articulo"].Value, dr.Cells["Codigo"].Value.ToString());
			
													oCmdActualiza.CommandText = sSQLDetalle;
													oCmdActualiza.ExecuteNonQuery();
												}

												#region Refrescar Detalle
												string sSQLActDetalle = string.Format("Exec ConsultaDetalleCompra {0}", (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLActDetalle;
												SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
												DataTable oDTDetalle = new DataTable();
												oDADetalle.Fill(oDTDetalle);
												this.ultraGrid1.DataSource = oDTDetalle;
												#endregion

												#endregion
												#endregion Graba Registro	

												#region valida dias
												if (Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count(*) From CompraDiasPagos Where idCompra = {0}",(int)this.txtNumIdCompra.Value)) != DiasCompra.uGridDetalle.Rows.Count)
												{
													string sSQLBorraDias = string.Format("Exec BorraCompraDiasPagos {0}", (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLBorraDias;
													oCmdActualiza.ExecuteNonQuery();
													bDiasCompra=true;
												}	
												#endregion valida dias
												/*3*/
									
												#region CompraDiasPagos
												if(bDiasCompra)
												{
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in DiasCompra.uGridDetalle.Rows.All)
													{
														int iPago = (int)dr.Cells["Pagos"].Value ;	
														int iDias = (int)dr.Cells["Dias"].Value;
														try
														{	

															dtFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT  DATEADD(DAY,{0},'{1}')",iDias,Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));
											
															string sSQLDiasPagos = string.Format("Exec GuardaCompraDiasPagos {0}, {1}, {2},'{3}','{4}'",(int)this.txtNumIdCompra.Value,iPago,iDias,dtFecha.ToString("yyyyMMdd"),dtFechaPago.ToString("yyyyMMdd"));
															oCmdActualiza.CommandText = sSQLDiasPagos;
															oCmdActualiza.ExecuteNonQuery();	
														}
														catch(Exception Exc)
														{								
															MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
														}
													}
												}
												#endregion CompraDiasPagos


												#region Guarda Retenciones
												DateTime dtFechaRet = (DateTime)Retenciones.dtFecha.Value;
												DateTime dtFechaCad = DateTime.Parse("01/01/2000");
												if (Retenciones.dtCaducidad.Value != System.DBNull.Value) dtFechaCad = (DateTime)Retenciones.dtCaducidad.Value;

												string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 1, 6, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
													(int) Retenciones.txtIdRetencion.Value, (int)this.txtNumIdCompra.Value, Retenciones.txtSerieFact.Text.ToString(), Retenciones.txtNumero.Text.ToString(), Retenciones.txtAutFactura.Text.ToString(), 
													dtFechaCad.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), Retenciones.txtNotas.Text.ToString(), (decimal)Retenciones.txtRetenidoIVA.Value, (decimal)Retenciones.txtRetenidoRenta.Value, 
													(decimal)Retenciones.txtRetenido.Value);
												oCmdActualiza.CommandText = sSQLGR;
												int idRetencion = (int)oCmdActualiza.ExecuteScalar();

												if (!bNuevo)
												{
													string sSQLValida = string.Format("Exec ValidaDetalleDeRetencion {0}", idRetencion);
													oCmdActualiza.CommandText = sSQLValida;
													oCmdActualiza.ExecuteNonQuery();
												}

												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Retenciones.ultraGrid1.Rows.All)
												{
													string sSQLGDR = string.Format("Exec GuardaDetalleRetenciones {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
														(int)dr.Cells["idDetalleRetencion"].Value, idRetencion, (bool)dr.Cells["IVARenta"].Value, (int)dr.Cells["idRetencionIVARenta"].Value, (decimal)dr.Cells["Base"].Value, (decimal)dr.Cells["Valor"].Value,
														(decimal)dr.Cells["Porcentaje"].Value);
													oCmdActualiza.CommandText = sSQLGDR;
													oCmdActualiza.ExecuteNonQuery();
												}

												string sSQLActRetenido = string.Format("Update Compra Set Retenido = {0} Where idCompra = {1}", (decimal)Retenciones.txtRetenido.Value, (int)this.txtNumIdCompra.Value);
												oCmdActualiza.CommandText = sSQLActRetenido;
												oCmdActualiza.ExecuteNonQuery();
												#endregion Guarda Retenciones

												#region Graba Google
								
												if(bNuevo)
												{
													if(txtDirectorioarchivo.Length > 10 )
													{
														NuevoGoogle (txtDirectorioarchivo);
													}
												}
												if (sUrlNube.Length > 10 )
												{
													string sSQLGoogle = string.Format("Exec GrabaUrlNube {0}, '{1}', '{2}'", (int)this.txtNumIdCompra.Value, sCarpeta, sUrlNube );
													oCmdActualiza.CommandText = sSQLGoogle;
													oCmdActualiza.ExecuteNonQuery();
												}

												#endregion Graba Google										

												#region RegistraBloqueoTransacciones

												if (!bNuevo)
												{
													string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLBloqueo;
													oCmdActualiza.ExecuteNonQuery();
												}
												else 
												{
													if ((int)this.txtNumidOrigen.Value > 0 && !bProforma)
													{
														string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
														oCmdActualiza.CommandText = sSQLBloqueoOrigen;
														oCmdActualiza.ExecuteNonQuery();
													}

													if ((int)this.txtNumidOrigen1.Value > 0)
													{
														string sSQLBloqueoOrigen1 = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
														oCmdActualiza.CommandText = sSQLBloqueoOrigen1;
														oCmdActualiza.ExecuteNonQuery();
													}
												}

												if (bProforma)
												{
													string sSQLProforma = string.Format("Update Proforma Set Estado = 3, FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME(), idCompra = {1} Where idProforma = {0}", 
														(int)this.txtNumidOrigen.Value, (int)this.txtNumIdCompra.Value);
													oCmdActualiza.CommandText = sSQLProforma;
													oCmdActualiza.ExecuteNonQuery();
												}
												#endregion RegistraBloqueoTransacciones

												oTransaction.Commit();

												#region Controles
												bNuevo = false;
												bEdicion = false;

												this.txtRuc.Enabled = false;
												this.txtSerieFact.Enabled = false;
												this.txtNumero.Enabled = false;
												this.txtAutFactura.Enabled = false;
												this.dtCaducidad.Enabled = false;
												this.dtFechaPeriodo.Enabled = false;
												this.dtFechaPeriodo.Enabled = false;
												this.txtDiasCredito1.Enabled = false;
												this.txtDiasCredito2.Enabled = false;
												this.txtDiasCredito3.Enabled = false;
												this.txtDiasCredito4.Enabled = false;
												this.txtDiasCredito5.Enabled = false;
												this.cmbTipoFormaPago.Enabled = false;
												this.cmbProyectoZOOM.Enabled = false;
												this.cmbImportacionNombre.Enabled = false;
												this.cmbCreditoTributario.Enabled = false;
												this.dtFecha.Enabled = false;
												//this.cmbFormaPago.Value = false;
												this.dtFechadeEmicion.Enabled = false;
												this.cmbFormaPago.Enabled = false;
										
			
												ModificaGrilla(false);
			
												this.txtNotas.Enabled = false;
						
												if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
												if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
												if (miAcceso.Editar) this.btnEditar.Enabled = true;
												DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
												DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
												if (miAcceso.EditarSoloDia)
												{
													if (dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString() )
													{
														this.btnEditar.Enabled = true;
													}
												}
												if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
												else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;			
												if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
												if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
												if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
												if (miAcceso.Retencion) this.btnRetencion.Enabled = true;
												this.btnGuardar.Enabled = false;
												this.btnTemporal.Enabled = false;
									

												if (sUrlNube.Length > 10 )
												{
													this.btnSubir.Visible =false;
													this.linkLabel1.Visible = true;
												}
						

												#endregion Controles						

												#region Estado
												if ((int)this.txtNumEstado.Value == 7) this.lblEstado.Text = "POR LIQUIDAR";
												if ((int)this.txtNumEstado.Value == 6) this.lblEstado.Text = "TEMPORAL";
												if ((int)this.txtNumEstado.Value == 8) this.lblEstado.Text = "LIQUIDADO";
												if ((int)this.txtNumEstado.Value == 14) this.lblEstado.Text = "LIQUIDADO";
												#endregion Estado

												#endregion Proceso

												if (bConsignacion) this.DialogResult = DialogResult.OK;
											}
											catch (Exception ex)
											{
												#region Exception
												try
												{
													oTransaction.Rollback();

													MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
												}
												catch (Exception ex2)
												{
													MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
												}
												#endregion Exception
											}
											finally
											{
												oConexion.Close();
											}
										}
										#endregion Caso Contrario
									}
									#endregion Proceso con retencion
								}
							}
						}
					}
				}
				#endregion Registro Con Retenciones
			}
			else
			{
				using (frmDiasCompra DiasCompra = new frmDiasCompra((int)this.txtNumIdCompra.Value,1))
				{
					if (DialogResult.OK == DiasCompra.ShowDialog())
					{
						if (DiasCompra.uGridDetalle.Rows.Count > 0)
						{	
							#region CasoContrario
							#region Valida Pagos Realizados
							decimal dPagado = 0.00m;
							decimal dAutorizado = 0.00m;
							decimal dPA = 0.00m;
							string sSQLValPagos = string.Format("Exec ValidaPagosCompras {0}", (int)this.txtNumIdCompra.Value);
							SqlDataReader drVP = Funcion.rEscalarSQL(cdsSeteoF, sSQLValPagos, true);
							drVP.Read();
							if (drVP.HasRows)
							{
								dPagado = drVP.GetDecimal(0);
								dAutorizado = drVP.GetDecimal(1);
							}
							drVP.Close();

							dPA = dPagado + dAutorizado;
							#endregion Valida Pagos Realizados

							if ((decimal)this.txtTotal.Value < dPA)
							{
								MessageBox.Show(string.Format("El Valor Pagado {0} más el Valor Autorizado {1} es Mayor al Valor Total de la Factura {2}", dPagado, dAutorizado, (decimal)this.txtTotal.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}
							else
							{
								using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

										#region Proceso

										#region Graba Registro
										#region Graba Maestro
										int idProyectoZOOM = 0;
										this.txtNumEstado.Value = 6;
										if (this.cmbProyectoZOOM.ActiveRow != null) idProyectoZOOM = (int)this.cmbProyectoZOOM.Value;

										//										string sSQLMaestro = string.Format("Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39},'{40}'",
										//											(int)this.txtNumIdCompra.Value, (int)this.txtNumidOrigen.Value, (int)this.txtNumidOrigen1.Value,
										//											(int)this.txtNumIdAsiento.Value, (int)this.txtIdSucursal.Value, (int)this.txtIdTipoFactura.Value, (int)this.cmbComprobante.Value, 
										//											idProyectoZOOM, (int)this.txtNumContadoCredito.Value, (int)this.txtNumEstado.Value, (int)this.txtNumIdCliente.Value, 
										//											(int)this.cmbBodega.Value, idProyecto, this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(),
										//											this.txtAutFactura.Text.ToString(), dtFecha.ToString("yyyyMMdd"), dtCaducidad.ToString("yyyyMMdd"),
										//											(int)this.cmbFormaPago.Value, (int)this.txtDiasCredito1.Value, (int)this.txtDiasCredito2.Value, (int)this.txtDiasCredito3.Value, IdTipoPago,
										//											IdImportacionNombre, this.txtComprobante.Text.ToString(), (bool)this.chkConsignacion.Checked, this.txtNotas.Text.ToString(),
										//											(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
										//											bFecCad, (int)this.txtIdTipoDePago.Value, (int)this.cmbCreditoTributario.Value, dtFechaRecepcion.ToString("yyyyMMdd HH:mm"), (int)this.txtDiasCredito3.Value, dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),(int)this.txtDiasCredito5.Value,dtFechadeEmicion.ToString("yyyyMMdd HH:mm"));
			
										string sSQLMaestro = string.Format(
											"Exec GrabaMaestroCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, {23}, '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, '{36}', {37}, '{38}', {39}, '{40}'",

											(int)this.txtNumIdCompra.Value,
											(int)this.txtNumidOrigen.Value,
											(int)this.txtNumidOrigen1.Value,
											(int)this.txtNumIdAsiento.Value,
											(int)this.txtIdSucursal.Value,
											(int)this.txtIdTipoFactura.Value,
											(int)this.cmbComprobante.Value,
											idProyectoZOOM,
											(int)this.txtNumContadoCredito.Value,
											(int)this.txtNumEstado.Value,
											(int)this.txtNumIdCliente.Value,
											(int)this.cmbBodega.Value,
											idProyecto,

											LimpiarValorSQL(this.txtSerieFact.Text),
											LimpiarValorSQL(this.txtNumero.Text),
											LimpiarValorSQL(this.txtAutFactura.Text),

											dtFecha.ToString("yyyyMMdd"),
											dtCaducidad.ToString("yyyyMMdd"),

											(int)this.cmbFormaPago.Value,
											(int)this.txtDiasCredito1.Value,
											(int)this.txtDiasCredito2.Value,
											(int)this.txtDiasCredito3.Value,
											IdTipoPago,
											IdImportacionNombre,

											LimpiarValorSQL(this.txtComprobante.Text),

											(bool)this.chkConsignacion.Checked,
											LimpiarValorSQL(this.txtNotas.Text),

											(decimal)this.txtIva0.Value,
											(decimal)this.txtDescIva0Total.Value,
											(decimal)this.txtIva.Value,
											(decimal)this.txtDescIvaTotal.Value,
											(decimal)this.txtIvaTotal.Value,
											(decimal)this.txtTotal.Value,

											bFecCad,
											(int)this.txtIdTipoDePago.Value,
											(int)this.cmbCreditoTributario.Value,

											dtFechaRecepcion.ToString("yyyyMMdd HH:mm"),
											(int)this.txtDiasCredito3.Value,
											dtFechaPEriodo.ToString("yyyyMMdd HH:mm"),
											(int)this.txtDiasCredito5.Value,
											dtFechadeEmicion.ToString("yyyyMMdd HH:mm")
											);

										oCmdActualiza.CommandText = sSQLMaestro;
										this.txtNumIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
										#endregion Graba Maestro

										#region Graba Detalle
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
										{	
											if (dr.Cells["idProyecto"].Value != System.DBNull.Value) IdProyectoDetalle = (int)dr.Cells["idProyecto"].Value; else IdProyectoDetalle = 0;
											if (dr.Cells["idImportacion"].Value != System.DBNull.Value) IdImportacionTipo = (int)dr.Cells["idImportacion"].Value; else IdImportacionTipo = 0;
				
											string sSQLDetalle = string.Format("Exec GrabaDetalleCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}', '{23}'", 
												(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, (int)this.txtNumIdCompra.Value, 
												(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, Math.Round((decimal)dr.Cells["DescuentoArt"].Value, 2), (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value, (decimal)dr.Cells["DescNC"].Value,
												IdProyectoDetalle, IdImportacionTipo, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["Signo"].Value, (int)dr.Cells["Tipo"].Value, 
												(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, (decimal)dr.Cells["Valores"].Value,
												(int)dr.Cells["idTipoGasto"].Value, dr.Cells["Articulo"].Value, dr.Cells["Codigo"].Value.ToString());
			
											oCmdActualiza.CommandText = sSQLDetalle;
											oCmdActualiza.ExecuteNonQuery();
										}

										#region Refrescar Detalle
										string sSQLActDetalle = string.Format("Exec ConsultaDetalleCompra {0}", (int)this.txtNumIdCompra.Value);
										oCmdActualiza.CommandText = sSQLActDetalle;
										SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
										DataTable oDTDetalle = new DataTable();
										oDADetalle.Fill(oDTDetalle);
										this.ultraGrid1.DataSource = oDTDetalle;							
										#endregion

										#endregion
										#endregion Graba Registro

										#region Graba Google
								
										if(bNuevo)
										{
											if(txtDirectorioarchivo.Length > 10 )
											{
												NuevoGoogle (txtDirectorioarchivo);
											}
										}
										if (sUrlNube.Length > 10 )
										{
											string sSQLGoogle = string.Format("Exec GrabaUrlNube {0}, '{1}', '{2}'", (int)this.txtNumIdCompra.Value, sCarpeta, sUrlNube );
											oCmdActualiza.CommandText = sSQLGoogle;
											oCmdActualiza.ExecuteNonQuery();
										}
							
										#endregion Graba Google


										#region valida dias
										if (Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count(*) From CompraDiasPagos Where idCompra = {0}",(int)this.txtNumIdCompra.Value)) != DiasCompra.uGridDetalle.Rows.Count)
										{
											string sSQLBorraDias = string.Format("Exec BorraCompraDiasPagos {0}", (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLBorraDias;
											oCmdActualiza.ExecuteNonQuery();
											bDiasCompra=true;
										}	
										#endregion valida dias
										/*4*/

										#region CompraDiasPagos
										if(bDiasCompra)
										{
											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in DiasCompra.uGridDetalle.Rows.All)
											{
												int iPago = (int)dr.Cells["Pagos"].Value ;	
												int iDias = (int)dr.Cells["Dias"].Value;
												try
												{	
													dtFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT  DATEADD(DAY,{0},'{1}')",iDias,Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));
											
													string sSQLDiasPagos = string.Format("Exec GuardaCompraDiasPagos {0}, {1}, {2},'{3}','{4}'",(int)this.txtNumIdCompra.Value,iPago,iDias,dtFecha.ToString("yyyyMMdd"),dtFechaPago.ToString("yyyyMMdd"));
													oCmdActualiza.CommandText = sSQLDiasPagos;
													oCmdActualiza.ExecuteNonQuery();	
												}
												catch(Exception Exc)
												{								
													MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
												}
											}
										}
										#endregion CompraDiasPagos

										#region RegistraBloqueoTransacciones

										if (!bNuevo)
										{
											string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLBloqueo;
											oCmdActualiza.ExecuteNonQuery();
										}
										else 
										{
											if ((int)this.txtNumidOrigen.Value > 0 && !bProforma)
											{
												string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen.Value);
												oCmdActualiza.CommandText = sSQLBloqueoOrigen;
												oCmdActualiza.ExecuteNonQuery();
											}

											if ((int)this.txtNumidOrigen1.Value > 0)
											{
												string sSQLBloqueoOrigen1 = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", (int)this.txtNumidOrigen1.Value);
												oCmdActualiza.CommandText = sSQLBloqueoOrigen1;
												oCmdActualiza.ExecuteNonQuery();
											}
										}

										if (bProforma)
										{
											string sSQLProforma = string.Format("Update Proforma Set Estado = 3, FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME(), idCompra = {1} Where idProforma = {0}", 
												(int)this.txtNumidOrigen.Value, (int)this.txtNumIdCompra.Value);
											oCmdActualiza.CommandText = sSQLProforma;
											oCmdActualiza.ExecuteNonQuery();
										}
										#endregion RegistraBloqueoTransacciones

										oTransaction.Commit();

										#region Controles
										bNuevo = false;
										bEdicion = false;

										this.txtRuc.Enabled = false;
										this.txtSerieFact.Enabled = false;
										this.txtNumero.Enabled = false;
										this.txtAutFactura.Enabled = false;
										this.dtCaducidad.Enabled = false;
										this.txtDiasCredito1.Enabled = false;
										this.txtDiasCredito2.Enabled = false;
										this.txtDiasCredito3.Enabled = false;
										this.txtDiasCredito4.Enabled = false;
										this.txtDiasCredito5.Enabled = false;
										//this.cmbFormaPago.Value = false;
										this.cmbTipoFormaPago.Enabled = false;
										this.cmbImportacionNombre.Enabled = false;
										this.cmbProyectoZOOM.Enabled = false;
										this.cmbCreditoTributario.Enabled = false;
										this.dtFecha.Enabled = false;
										this.dtFechadeEmicion.Enabled = false;
										this.cmbFormaPago.Enabled = false;
								
			
										ModificaGrilla(false);
			
										this.txtNotas.Enabled = false;
						
										if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
										if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
										if (miAcceso.Editar) this.btnEditar.Enabled = true;
										DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
										DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
										if (miAcceso.EditarSoloDia)
										{
											if (dtFechaFac.ToShortDateString() == dtFechaHoy.ToShortDateString() )
											{
												this.btnEditar.Enabled = true;
											}
										}
										if (!bImpreso && miAcceso.Imprimir) this.btnImprimir.Enabled = true;
										else if (bImpreso && miAcceso.ReImprimir) this.btnImprimir.Enabled = true;			
										if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
										if (miAcceso.BBuscarOrdenCompra) this.btnBuscarOrdenDeCompra.Enabled = true;
										if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
										if (miAcceso.Retencion) this.btnRetencion.Enabled = true;
										this.btnGuardar.Enabled = false;
										this.btnTemporal.Enabled = false;
								
										this.btnCancelar.Enabled = true;

										if (sUrlNube.Length > 10 )
										{
											this.btnSubir.Visible =false;
											this.linkLabel1.Visible = true;
										}

										#endregion Controles						

								
										#region Estado
										if ((int)this.txtNumEstado.Value == 7) this.lblEstado.Text = "POR LIQUIDAR";
										if ((int)this.txtNumEstado.Value == 6) this.lblEstado.Text = "TEMPORAL";
										if ((int)this.txtNumEstado.Value == 8) this.lblEstado.Text = "LIQUIDADO";
										if ((int)this.txtNumEstado.Value == 14) this.lblEstado.Text = "LIQUIDADO";
										#endregion Estado

										#endregion Proceso

										if (bConsignacion) this.DialogResult = DialogResult.OK;
									}						
									catch (Exception ex)
									{
										#region Exception
										try
										{
											oTransaction.Rollback();

											MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
										}
										catch (Exception ex2)
										{
											MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
										}
										#endregion Exception
									}
									finally
									{
										oConexion.Close();
									}
								}
							}
							#endregion CasoContrario
						}
					}
				}
			}
			#endregion Graba Registro
			
			
		}

		private void mnuDistribucion_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void btnDias_Click(object sender, System.EventArgs e)
		{
			using (frmDiasCompra DiasCompra = new frmDiasCompra((int)this.txtNumIdCompra.Value,2))
			{
				if (DialogResult.OK == DiasCompra.ShowDialog())
				{
				}
			}
		}

		private int ObtenerIndiceColumna(SqlDataReader dr, string nombreColumna)
		{
			for (int i = 0; i < dr.FieldCount; i++)
			{
				if (string.Compare(dr.GetName(i), nombreColumna, true) == 0)
					return i;
			}

			return -1;
		}

		private bool ValidaPresupuestoCuentaGrid()
		{
			try
			{
				ArrayList mensajes = new ArrayList();
				DateTime dtFechaValida = (DateTime)this.dtFecha.Value;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					if (dr == null) continue;
					if (!dr.Cells.Exists("idProyecto") || dr.Cells["idProyecto"].Value == System.DBNull.Value) continue;
					if (!dr.Cells.Exists("Codigo") || dr.Cells["Codigo"].Value == System.DBNull.Value) continue;

					int idProyectoValida = (int)dr.Cells["idProyecto"].Value;
					string  sCodigoValida = (string)dr.Cells["Codigo"].Value;

					decimal dNuevoValor = 0.00m;
					if (dr.Cells.Exists("SubTotal") && dr.Cells["SubTotal"].Value != System.DBNull.Value)
						dNuevoValor = Math.Round(Convert.ToDecimal(dr.Cells["SubTotal"].Value), 2);

					string sSQLValidasms = string.Format(
						"Exec ValidaCuentaDentroPresupuesto '{0}', {1}, '{2}', {3}",
						dtFechaValida.ToString("yyyyMMdd"),
						idProyectoValida,
						sCodigoValida,
						dNuevoValor.ToString("0.00", us));

					SqlDataReader drValida = Funcion.rEscalarSQL(cdsSeteoF, sSQLValidasms, true);
					int idxMensaje = ObtenerIndiceColumna(drValida, "Mensaje");

					while (drValida.Read())
					{
						string mensaje = "";
						if (idxMensaje >= 0 && drValida[idxMensaje] != System.DBNull.Value)
							mensaje = drValida[idxMensaje].ToString().Trim();
						else if (drValida.FieldCount > 0 && drValida[0] != System.DBNull.Value)
							mensaje = drValida[0].ToString().Trim();

						if (mensaje.Trim().Length > 0 && !mensajes.Contains(mensaje))
							mensajes.Add(mensaje);
					}

					drValida.Close();
				}

				if (mensajes.Count > 0)
				{
					string sMensajeFinal = "Se encontraron validaciones de presupuesto:\n\n";
					for (int i = 0; i < mensajes.Count; i++)
					{
						sMensajeFinal += "- " + mensajes[i].ToString() + "\n\n";
					}

					MessageBox.Show(sMensajeFinal, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Question);
					return false;
				}

				return true;
			}
			catch (Exception Exc)
			{
				MessageBox.Show(string.Format("Error al validar presupuesto por cuenta: {0}", Exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}

		private int ValidaPresupuestoCuentaGrid10()
		{
			try
			{
				DateTime dtFechaValida = (DateTime)this.dtFecha.Value;
				bool bTieneMensajes = false;
				int idCompra = 0;
				if (this.txtNumIdCompra.Value != System.DBNull.Value)
					idCompra = (int)this.txtNumIdCompra.Value;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					if (dr == null) continue;
					if (!dr.Cells.Exists("idProyecto") || dr.Cells["idProyecto"].Value == System.DBNull.Value) continue;
					if (!dr.Cells.Exists("Codigo") || dr.Cells["Codigo"].Value == System.DBNull.Value) continue;
					if (!dr.Cells.Exists("idDetCompra") || dr.Cells["idDetCompra"].Value == System.DBNull.Value) continue;

					int idProyectoValida = (int)dr.Cells["idProyecto"].Value;
					string sCodigoValida = dr.Cells["Codigo"].Value.ToString().Trim();
					int idDetCompra = (int)dr.Cells["idDetCompra"].Value;

					decimal dNuevoValor = 0.00m;
					if (dr.Cells.Exists("SubTotal") && dr.Cells["SubTotal"].Value != System.DBNull.Value)
						dNuevoValor = Math.Round(Convert.ToDecimal(dr.Cells["SubTotal"].Value), 2);

					string sSQLValida = string.Format(
						"Exec ValidaCuentaDentroPresupuesto '{0}', {1}, '{2}', {3}",
						dtFechaValida.ToString("yyyyMMdd"),
						idProyectoValida,
						sCodigoValida,
						dNuevoValor.ToString("0.00", us));

					SqlDataReader drValida = Funcion.rEscalarSQL(cdsSeteoF, sSQLValida, true);
					int idxMensaje = ObtenerIndiceColumna(drValida, "Mensaje");
					ArrayList mensajesFila = new ArrayList();

					while (drValida.Read())
					{
						string mensaje = "";
						if (idxMensaje >= 0 && drValida[idxMensaje] != System.DBNull.Value)
							mensaje = drValida[idxMensaje].ToString().Trim();
						else if (drValida.FieldCount > 0 && drValida[0] != System.DBNull.Value)
							mensaje = drValida[0].ToString().Trim();

						if (mensaje.Length > 0)
							mensajesFila.Add(mensaje);
					}

					drValida.Close();

					for (int i = 0; i < mensajesFila.Count; i++)
					{
						string sSQLCorreo = string.Format(
							"Exec GuardaPresupuestoGastosAnualEmail '{0}', '{1}', '{2}', '{3}', {4}, '{5}', {6}",
							LimpiarValorSQL(this.txtRuc.Text),
							LimpiarValorSQL(this.txtNombre.Text),
							LimpiarValorSQL(this.txtSerieFact.Text),
							LimpiarValorSQL(this.txtNumero.Text),
							idCompra,
							LimpiarValorSQL(mensajesFila[i].ToString()),
							idDetCompra);

						Funcion.EjecutaSQL(cdsSeteoF, sSQLCorreo, true);
						bTieneMensajes = true;
					}
				}

				return bTieneMensajes ? 0 : 1;
			}
			catch
			{
				return 0;
			}
		}

		private void CargarVentaVerdeEnDetalle(DataTable dtProductos)
		{
			bool bActualizaAnterior = bActualiza;

			try
			{
				bActualiza = false;

				foreach (DataRow dr in dtProductos.Rows)
				{
					UltraGridRow row = this.ultraGrid1.DisplayLayout.Bands[0].AddNew();

					decimal costo = 0;
					if (dr.Table.Columns.Contains("Costo") &&
						dr["Costo"] != null &&
						dr["Costo"] != System.DBNull.Value)
					{
						costo = Convert.ToDecimal(dr["Costo"]);
					}

					decimal cantidad = 1;
					if (dr.Table.Columns.Contains("Cantidad") &&
						dr["Cantidad"] != null &&
						dr["Cantidad"] != System.DBNull.Value)
					{
						cantidad = Convert.ToDecimal(dr["Cantidad"]);
					}

					decimal subtotal = cantidad * costo;

					int idDetCompraOrigenVV = 0;
					if (dr.Table.Columns.Contains("idDetCompraOrigenVV") &&
						dr["idDetCompraOrigenVV"] != null &&
						dr["idDetCompraOrigenVV"] != System.DBNull.Value)
					{
						idDetCompraOrigenVV = Convert.ToInt32(dr["idDetCompraOrigenVV"]);
					}

					if (row.Cells.Exists("idDetCompraVV"))
						row.Cells["idDetCompraVV"].Value = idDetCompraOrigenVV;

					row.Cells["idOrigen"].Value = 0;

					int idArticulo = Convert.ToInt32(dr["idArticulo"]);

					CargarInformacionArticuloCompraEnFila(
						row,
						idArticulo,
						cantidad,
						costo
						);

					row.Cells["EsVentaVerde"].Value = true;
				}

				this.ultraGrid1.UpdateData();
				cmdTotal();
			}
			finally
			{
				bActualiza = bActualizaAnterior;
			}
		}
		private void MarcarVentaVerdeLiquidada(SqlCommand oCmdActualiza)
		{
			int idClienteProveedor = 0;
			bool existeVentaVerde = false;

			if (this.txtNumIdCliente.Value != null && this.txtNumIdCliente.Value != DBNull.Value)
				idClienteProveedor = Convert.ToInt32(this.txtNumIdCliente.Value);

			if (idClienteProveedor <= 0)
				throw new Exception("No se pudo obtener el proveedor para procesar Venta Verde.");

			// 1. Verificar si esta compra tiene al menos un detalle Venta Verde
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.ultraGrid1.Rows.All)
			{
				int idDetCompraVV = ObtenerIdDetCompraVV(row);

				if (idDetCompraVV > 0)
				{
					existeVentaVerde = true;
					break;
				}
			}

			if (!existeVentaVerde)
				return;

			// 2. Consultar porcentaje de comisión desde Cliente.DescuentoContado
			decimal porcentajeComision = ObtenerDescuentoContadoCliente(oCmdActualiza, idClienteProveedor);

			int idFacturaComision = 0;

			// 3. Si tiene porcentaje, generar UNA sola factura de venta por comisión
			if (porcentajeComision > 0)
			{
				decimal totalCompra = Convert.ToDecimal(this.txtTotal.Value);

				idFacturaComision = GenerarFacturaComisionVentaVerde(
					oCmdActualiza,
					idClienteProveedor,
					Convert.ToInt32(this.txtNumIdCompra.Value),
					totalCompra,
					porcentajeComision
					);
			}

			// 4. Marcar todos los detalles Venta Verde como liquidados
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.ultraGrid1.Rows.All)
			{
				int idDetCompraVV = ObtenerIdDetCompraVV(row);

				if (idDetCompraVV <= 0)
					continue;

				oCmdActualiza.CommandText = string.Format(
					"Exec ActualizaEstadoDetCompraVV {0}",
					idDetCompraVV
					);

				oCmdActualiza.ExecuteNonQuery();
			}
		}

		private int ObtenerIdDetCompraVV(Infragistics.Win.UltraWinGrid.UltraGridRow row)
		{
			if (row == null) return 0;
			if (!row.Cells.Exists("idDetCompraVV")) return 0;

			object valor = row.Cells["idDetCompraVV"].Value;

			if (valor == null || valor == DBNull.Value) return 0;

			try
			{
				return Convert.ToInt32(valor);
			}
			catch
			{
				return 0;
			}
		}

		
		private decimal ObtenerDescuentoContadoCliente(SqlCommand oCmdActualiza, int idCliente)
		{
			oCmdActualiza.CommandText = string.Format(
				"Select IsNull(DescuentoContado, 0) From Cliente Where idCliente = {0}",
				idCliente
				);

			object valor = oCmdActualiza.ExecuteScalar();

			if (valor == null || valor == DBNull.Value)
				return 0.00m;

			return Convert.ToDecimal(valor);
		}

		
		private int GenerarFacturaComisionVentaVerde(
			SqlCommand oCmdActualiza,
			int idClienteProveedor,
			int idCompraLiquidacion,
			decimal totalCompra,
			decimal porcentajeComision
			)
		{
			oCmdActualiza.CommandText = string.Format(
				"Exec GeneraFacturaComisionVentaVerde {0}, {1}, {2}, {3}",
				idClienteProveedor,
				idCompraLiquidacion,
				totalCompra.ToString(System.Globalization.CultureInfo.InvariantCulture),
				porcentajeComision.ToString(System.Globalization.CultureInfo.InvariantCulture)
				);

			object resultado = oCmdActualiza.ExecuteScalar();

			if (resultado == null || resultado == DBNull.Value)
				throw new Exception("No se generó la factura de comisión Venta Verde.");

			return Convert.ToInt32(resultado);
		}


		private void CargarInformacionArticuloCompraEnFila(
			UltraGridRow row,
			int idArticulo,
			decimal cantidad,
			decimal precio
			)
		{
			string stSelect = string.Format(
				"Exec CargaInformArticuloCompra {0}, {1}, 4, '{2}', 1",
				idArticulo,
				Convert.ToInt32(this.txtIdGrupoCliente.Value),
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")
				);

			SqlDataReader dr = null;

			try
			{
				dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);

				if (!dr.Read())
					throw new Exception("No se encontró información del artículo para cargar la compra.");

				row.Cells["Codigo"].Value = dr.GetString(0);
				row.Cells["Articulo"].Value = dr.GetString(1);
				row.Cells["idUnidad"].Value = dr.GetInt32(2);
				row.Cells["Cantidad"].Value = cantidad;
				row.Cells["Precio"].Value = precio;
				row.Cells["Impuesto"].Value = dr.GetDecimal(4);
				row.Cells["idArticulo"].Value = idArticulo;
			}
			finally
			{
				if (dr != null)
					dr.Close();
			}
		}


		///
		private void mnuAgenciasI_Click(object sender, System.EventArgs e)
		{
			Distribucion(4);
		}

		private void mnuAdministrativosI_Click(object sender, System.EventArgs e)
		{
			Distribucion(5);
		}

		private void mnuGeneralI_Click(object sender, System.EventArgs e)
		{
			Distribucion(6);
		}

		private bool ValidaEstadoDocumento()
		{
			bool requiereDocumento = false;

			int idDB = MenuLatinium.stIdDB;
			int idGrupoCliente = (int)this.txtIdGrupoCliente.Value;
			int tipoComprobante = (int)this.cmbComprobante.Value;
			decimal total = (decimal)this.txtTotal.Value;
			int tipoRuc = (int)this.txtIdTipoRuc.Value;

			// ------------------------------------------
			// 1. Determinar si se requiere subir documento
			// ------------------------------------------
			if (idDB != 1 && idDB != 33 && idDB != 3)
			{
				if (total >= 1000m)
					requiereDocumento = true;
			}
			else
			{
				if (idGrupoCliente != 5)
				{
					requiereDocumento = true;
				}
				else
				{
					if (tipoComprobante != 3 && total > 100m)
						requiereDocumento = true;

					if (tipoComprobante == 3 && tipoRuc == 6)
						requiereDocumento = true;
				}
			}

			// ------------------------------------------
			// 2. Validar si ya existe archivo seleccionado
			// ------------------------------------------
			bool hayArchivo = (sUrlNube != null && sUrlNube != "" && sUrlNube.Length > 10);

			// Si ya existe archivo o NO se requiere documento no pedir más
			if (hayArchivo || requiereDocumento == false)
				return true;

			// ------------------------------------------
			// 3. Intentar subir archivo
			// ------------------------------------------
			btnSubir_Click(null, EventArgs.Empty);

			if (!bNuevo)
			{
				// Modo editar: debe existir sUrlNube después del intento de subida
				if (sUrlNube == null || sUrlNube.Length < 10)
				{
					MessageBox.Show("Seleccione y suba el documento para continuar.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			else
			{
				// Modo nuevo: debe existir carpeta generada al subir
				if (sCarpeta == null || sCarpeta.Length < 2)
				{
					MessageBox.Show("Seleccione y suba el documento para continuar.",
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}

			// ------------------------------------------
			// 4. Consumir API
			// ------------------------------------------
			string url =

				MenuLatinium.stApiFacturasCompra +
				MenuLatinium.stDirFacturacion + "/" +
				this.txtRuc.Text + "/" +
				sCarpeta;

			string jsonRespuesta = Funcion.ConsumirAPI(url);

			if (jsonRespuesta == "Error")
			{
				MessageBox.Show("Error en la subida de Imágenes.\n\nComunícate con el Área de Desarrollo.",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			// ------------------------------------------
			// 5. Construir URL final
			// ------------------------------------------
			sUrlNube =
				"https://storage.googleapis.com/sparta_bucket/Compras/" +
				MenuLatinium.stDirFacturacion + "/" +
				this.txtRuc.Text + "/" +
				sCarpeta;

			return true;
		}

		private void cmbTiposDeGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}


}

