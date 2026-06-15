using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmOrdenesDeCompra.
	/// </summary>
	public class frmOrdenesDeCompra : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaLlegada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label label52;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnOrdenDeCompraMixta;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbImportacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCredTributario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnEMail;
		private System.Windows.Forms.Button btnImportacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProyecto;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		int IdOC = 0;
		DateTime hl;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dthorallegada;
		int iOrigen = 0;
		string SColumna="";
		public frmOrdenesDeCompra()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmOrdenesDeCompra(int idOC)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
		
			IdOC = idOC;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmOrdenesDeCompra(int idOC, int IOrigen, DateTime Hl, string sColumna)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
		
			InitializeComponent();

			IdOC = idOC;
			iOrigen = IOrigen;
			hl=Hl ;
			SColumna= sColumna;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmOrdenesDeCompra));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantAereo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valores");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Courier");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Arancel");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescNC");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantAereo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Courier");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Arancel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescNC");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.dtFechaLlegada = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.cmbImportacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.label52 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnOrdenDeCompraMixta = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCredTributario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEMail = new System.Windows.Forms.Button();
			this.btnImportacion = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.dthorallegada = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dthorallegada)).BeginInit();
			this.SuspendLayout();
			// 
			// dtFechaLlegada
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaLlegada.Appearance = appearance1;
			this.dtFechaLlegada.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaLlegada.DateButtons.Add(dateButton1);
			this.dtFechaLlegada.Enabled = false;
			this.dtFechaLlegada.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaLlegada.Location = new System.Drawing.Point(576, 40);
			this.dtFechaLlegada.Name = "dtFechaLlegada";
			this.dtFechaLlegada.NonAutoSizeHeight = 23;
			this.dtFechaLlegada.Size = new System.Drawing.Size(136, 21);
			this.dtFechaLlegada.SpinButtonsVisible = true;
			this.dtFechaLlegada.TabIndex = 281;
			this.dtFechaLlegada.Value = ((object)(resources.GetObject("dtFechaLlegada.Value")));
			this.dtFechaLlegada.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaLlegada_Validating);
			this.dtFechaLlegada.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaLlegada_BeforeDropDown);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(872, 456);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 280;
			this.txtIdCompra.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(472, 69);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 26);
			this.lblUsuario.TabIndex = 279;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbImportacion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbImportacion.Appearance = appearance2;
			this.cmbImportacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbImportacion.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 270;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbImportacion.DisplayMember = "Nombre";
			this.cmbImportacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbImportacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbImportacion.Enabled = false;
			this.cmbImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbImportacion.Location = new System.Drawing.Point(112, 72);
			this.cmbImportacion.Name = "cmbImportacion";
			this.cmbImportacion.Size = new System.Drawing.Size(260, 21);
			this.cmbImportacion.TabIndex = 277;
			this.cmbImportacion.ValueMember = "idImportacionNombre";
			this.cmbImportacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbImportacion_KeyDown);
			this.cmbImportacion.Validating += new System.ComponentModel.CancelEventHandler(this.cmbImportacion_Validating);
			this.cmbImportacion.ValueChanged += new System.EventHandler(this.cmbImportacion_ValueChanged);
			this.cmbImportacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbImportacion_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// chkConsignacion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance3;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.CausesValidation = false;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(592, 7);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 276;
			this.chkConsignacion.Text = "Consignación";
			this.chkConsignacion.Visible = false;
			this.chkConsignacion.CheckedChanged += new System.EventHandler(this.chkConsignacion_CheckedChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(736, 69);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 275;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(832, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 273;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(752, 414);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 272;
			this.label1.Text = "TOTAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 10);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(38, 16);
			this.lblIdentificacion.TabIndex = 269;
			this.lblIdentificacion.Text = "Código";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.Location = new System.Drawing.Point(472, 42);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(86, 16);
			this.label52.TabIndex = 268;
			this.label52.Text = "Inicio Recepción";
			this.label52.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(752, 339);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(58, 16);
			this.lblSubTotal.TabIndex = 259;
			this.lblSubTotal.Text = "SUBTOTAL";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(200, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 258;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 256;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(752, 387);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 255;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(752, 363);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(67, 16);
			this.lblDescuento.TabIndex = 254;
			this.lblDescuento.Text = "DESCUENTO";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(984, 312);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 253;
			this.lblIvaSubt.Text = "IVA";
			this.lblIvaSubt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(880, 312);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 252;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(200, 42);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 257;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance4;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(64, 7);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 249;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(264, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 21);
			this.cmbBodega.TabIndex = 267;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtNombre
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance5;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(264, 7);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(312, 22);
			this.txtNombre.TabIndex = 248;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance6;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(832, 408);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 28);
			this.txtTotal.TabIndex = 266;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance7;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(928, 384);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 265;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance8;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(832, 360);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 264;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance9;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(928, 360);
			this.txtDescIvaTotal.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 263;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance10;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(928, 336);
			this.txtIva.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 262;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance11;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(832, 336);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 261;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance12;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(896, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 247;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 15;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Artículo";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 194;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Unidad";
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn10.Width = 53;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 185;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 340;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance14;
			ultraGridColumn13.Format = "#,##0";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.MaskInput = "";
			ultraGridColumn13.NullText = "0";
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance15;
			ultraGridColumn14.Format = "#,##0";
			ultraGridColumn14.Header.Caption = "Avión";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.MaskInput = "";
			ultraGridColumn14.NullText = "";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 70;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance16;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 84;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance17;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "IVA";
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 50;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance18;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Descuento";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 14;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance19;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% Desc";
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance20;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 100;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance21;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 25;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Centro de costo";
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn21.Width = 150;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance22;
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.MaskInput = "nnnnnnnnn";
			ultraGridColumn23.Width = 14;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 18;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 19;
			ultraGridColumn26.Hidden = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance23;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance24;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 21;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance25;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 22;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Hidden = true;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance26;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Dscto Prom";
			ultraGridColumn31.Header.VisiblePosition = 24;
			ultraGridColumn31.Width = 90;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance27;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 8.5F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance30;
			appearance31.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance31.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1000, 200);
			this.ultraGrid1.TabIndex = 251;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
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
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
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
			ultraDataColumn23.DataType = typeof(bool);
			ultraDataColumn23.DefaultValue = false;
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn28.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn28});
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance34;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 344);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(550, 88);
			this.txtNotas.TabIndex = 246;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 320);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 283;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 440);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1016, 8);
			this.gbBotones.TabIndex = 356;
			this.gbBotones.TabStop = false;
			// 
			// btnOrdenDeCompraMixta
			// 
			this.btnOrdenDeCompraMixta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnOrdenDeCompraMixta.CausesValidation = false;
			this.btnOrdenDeCompraMixta.Enabled = false;
			this.btnOrdenDeCompraMixta.Image = ((System.Drawing.Image)(resources.GetObject("btnOrdenDeCompraMixta.Image")));
			this.btnOrdenDeCompraMixta.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnOrdenDeCompraMixta.Location = new System.Drawing.Point(482, 456);
			this.btnOrdenDeCompraMixta.Name = "btnOrdenDeCompraMixta";
			this.btnOrdenDeCompraMixta.Size = new System.Drawing.Size(142, 23);
			this.btnOrdenDeCompraMixta.TabIndex = 350;
			this.btnOrdenDeCompraMixta.Text = "&Orden de Compra Mixta";
			this.btnOrdenDeCompraMixta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnOrdenDeCompraMixta.Click += new System.EventHandler(this.btnOrdenDeCompraMixta_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 456);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 345;
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
			this.btnConsultar.Location = new System.Drawing.Point(83, 456);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 354;
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
			this.btnEditar.Location = new System.Drawing.Point(164, 456);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 353;
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
			this.btnImprimir.Location = new System.Drawing.Point(243, 456);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 347;
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
			this.btnGuardar.Location = new System.Drawing.Point(403, 456);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 346;
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
			this.btnAnular.Location = new System.Drawing.Point(627, 456);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 355;
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
			this.btnCancelar.Location = new System.Drawing.Point(706, 456);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 348;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ForeColor = System.Drawing.Color.Black;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(781, 456);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 349;
			this.btnSalir.Text = "&Salir";
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(888, 456);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 358;
			this.txtEstado.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(904, 456);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdCliente.TabIndex = 357;
			this.txtIdCliente.Visible = false;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(936, 456);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdGrupoCliente.TabIndex = 362;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(920, 456);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoRuc.TabIndex = 361;
			this.txtIdTipoRuc.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(952, 456);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 360;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtContadoCredito
			// 
			this.txtContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContadoCredito.Enabled = false;
			this.txtContadoCredito.Location = new System.Drawing.Point(968, 456);
			this.txtContadoCredito.Name = "txtContadoCredito";
			this.txtContadoCredito.PromptChar = ' ';
			this.txtContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtContadoCredito.TabIndex = 363;
			this.txtContadoCredito.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(984, 456);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 364;
			this.txtIdComprobante.Visible = false;
			// 
			// txtIdCredTributario
			// 
			this.txtIdCredTributario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCredTributario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCredTributario.Enabled = false;
			this.txtIdCredTributario.Location = new System.Drawing.Point(1000, 456);
			this.txtIdCredTributario.Name = "txtIdCredTributario";
			this.txtIdCredTributario.PromptChar = ' ';
			this.txtIdCredTributario.Size = new System.Drawing.Size(16, 22);
			this.txtIdCredTributario.TabIndex = 365;
			this.txtIdCredTributario.Visible = false;
			// 
			// btnEMail
			// 
			this.btnEMail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEMail.CausesValidation = false;
			this.btnEMail.Enabled = false;
			this.btnEMail.Image = ((System.Drawing.Image)(resources.GetObject("btnEMail.Image")));
			this.btnEMail.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEMail.Location = new System.Drawing.Point(323, 456);
			this.btnEMail.Name = "btnEMail";
			this.btnEMail.Size = new System.Drawing.Size(77, 23);
			this.btnEMail.TabIndex = 366;
			this.btnEMail.Text = "&E-Mail";
			this.btnEMail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEMail.Click += new System.EventHandler(this.btnEMail_Click);
			// 
			// btnImportacion
			// 
			this.btnImportacion.CausesValidation = false;
			this.btnImportacion.Image = ((System.Drawing.Image)(resources.GetObject("btnImportacion.Image")));
			this.btnImportacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportacion.Location = new System.Drawing.Point(8, 71);
			this.btnImportacion.Name = "btnImportacion";
			this.btnImportacion.Size = new System.Drawing.Size(88, 23);
			this.btnImportacion.TabIndex = 367;
			this.btnImportacion.Text = "&Importación";
			this.btnImportacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportacion.Click += new System.EventHandler(this.btnImportacion_Click);
			// 
			// dtFecha
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance35;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 370;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtIdProyecto
			// 
			this.txtIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdProyecto.Enabled = false;
			this.txtIdProyecto.Location = new System.Drawing.Point(856, 456);
			this.txtIdProyecto.Name = "txtIdProyecto";
			this.txtIdProyecto.PromptChar = ' ';
			this.txtIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtIdProyecto.TabIndex = 371;
			this.txtIdProyecto.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.label4.ForeColor = System.Drawing.Color.Black;
			this.label4.Location = new System.Drawing.Point(736, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(73, 16);
			this.label4.TabIndex = 372;
			this.label4.Text = "Fin Recepción";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dthorallegada
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dthorallegada.Appearance = appearance36;
			this.dthorallegada.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dthorallegada.DateButtons.Add(dateButton3);
			this.dthorallegada.Enabled = false;
			this.dthorallegada.Format = "dd/MM/yyyy HH:mm";
			this.dthorallegada.Location = new System.Drawing.Point(832, 40);
			this.dthorallegada.Name = "dthorallegada";
			this.dthorallegada.NonAutoSizeHeight = 23;
			this.dthorallegada.Size = new System.Drawing.Size(136, 21);
			this.dthorallegada.SpinButtonsVisible = true;
			this.dthorallegada.TabIndex = 374;
			this.dthorallegada.Value = ((object)(resources.GetObject("dthorallegada.Value")));
			// 
			// frmOrdenesDeCompra
			// 
			this.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Appearance.Options.UseFont = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1016, 486);
			this.Controls.Add(this.dthorallegada);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.label52);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtIdProyecto);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnImportacion);
			this.Controls.Add(this.btnEMail);
			this.Controls.Add(this.txtIdCredTributario);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtContadoCredito);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnOrdenDeCompraMixta);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.dtFechaLlegada);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.cmbImportacion);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNotas);
			this.KeyPreview = true;
			this.Name = "frmOrdenesDeCompra";
			this.Text = "Ordenes de Compra";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmOrdenesDeCompra_KeyDown);
			this.Load += new System.EventHandler(this.frmOrdenesDeCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dthorallegada)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		bool bRastreo = false;
		int IdSigno = 1;
		private C1DataRow drSeteoF;
		private bool bActualiza = true;
		decimal dPorcIva = 0.00m;
		int idBloqueaTransacciones = 0;
		int idProyecto = 0;
		bool bCarga;
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void frmOrdenesDeCompra_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0602");

			bCarga=true;

			if (!Funcion.Permiso("102", cdsSeteoF))
			{				
				MessageBox.Show("No Puede Ingresar a Ordenes de Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(this.Close));				

				return;
			}
				
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'ORDEN DE COMPRA'", true);

			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");

			IdSigno = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 2");
				
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 2, 3", true);
				
			FuncionesProcedimientos.ListaImportacionesActivas(this.cmbImportacion, 0, cdsSeteoF);

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
//			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);

			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			if (IdOC > 0) this.ConsultaMaestro(IdOC);
	
			if (iOrigen == 1)
			{				
				if (IdOC > 0 && miAcceso.Editar) this.btnEditar_Click(sender, e); else if (IdOC == 0 && miAcceso.Nuevo) this.btnNuevo_Click(sender, e);

				if (bNuevo)
				{
					this.dthorallegada.Value = hl.AddMinutes(30);
					this.dtFechaLlegada.Value = hl;
					this.dtFechaLlegada.Enabled = false;				
				}

				if (bEdicion)				
					if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)					
						this.dtFechaLlegada.Value=hl;						
			}

			bCarga = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bNuevo = false;
			bEdicion = false;

			this.txtIdCompra.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdProyecto.Value = 0;
			this.txtIdGrupoCliente.Value = 0;
			this.txtIdTipoRuc.Value = 0;
			this.txtContadoCredito.Value = 0;
			this.txtIdComprobante.Value = 0;
			this.txtIdCredTributario.Value = 0;
			
			this.chkConsignacion.Checked = false;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.cmbBodega.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");
			this.dtFechaLlegada.Value = System.DBNull.Value;
			this.dthorallegada.Value =  System.DBNull.Value;
			this.txtNumero.Text = "";
			this.cmbImportacion.Value = System.DBNull.Value;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenCompra 0");

			this.txtNotas.Text = "";
			
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			bAnulado = false;
			bEliminaAlValidar = false;

			this.chkConsignacion.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaLlegada.Enabled = false;		
			this.dthorallegada.Enabled = false;
			this.txtNumero.Enabled = false;
			this.cmbImportacion.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			this.txtNotas.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnOrdenDeCompraMixta.Enabled = false;
			this.btnEMail.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Text = "Anular";

			this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = true;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = true;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Courier"].Hidden = true;

			idBloqueaTransacciones = 0;
			dPorcIva = 0.00m;

			if (iOrigen == 1) this.Close();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (iOrigen != 1) this.btnCancelar_Click(sender, e);

			this.txtEstado.Value = 1;						
			
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			if (iOrigen != 1)
			{
				this.dtFechaLlegada.Enabled=true;
				FuncionesProcedimientos.FechasDocumentos(this.dtFechaLlegada);//this.dtFechaLlegada.Value = DateTime.Today;
			}
			//			this.dthorallegada.Enabled=true;
						
			this.chkConsignacion.Enabled = true;
			this.txtRuc.Enabled = true;
				
			this.cmbBodega.Enabled = true;
			this.txtNotas.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnOrdenDeCompraMixta.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.txtRuc.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(true, 2);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					if (miAcceso.CambiarFecha) this.dtFecha.Focus();	
					else this.dtFechaLlegada.Focus();
				}				
			}

			if (e.KeyCode == Keys.Enter)
			{
				if (miAcceso.CambiarFecha)				this.dtFecha.Focus();				
				else this.dtFechaLlegada.Focus();
			}
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(true, 2, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.txtIdTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[2] != null) this.txtIdGrupoCliente.Value = sParametros[2];
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
			if (sParametros[6] != null) this.txtContadoCredito.Value = sParametros[6];
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				CargaCliente();

				this.chkConsignacion.Checked = false;
				this.chkConsignacion.Enabled = false;
				this.dtFechaLlegada.Enabled = false;		
				this.dthorallegada.Enabled = false;	
				this.cmbImportacion.Enabled = false;
	
				if ((int)this.txtIdCliente.Value > 0)
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

					string sSQLMail = string.Format("Select Count(*) From ClienteContacto Where EnvioMail = 1 And idCliente = {0}", (int)this.txtIdCliente.Value);
					int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLMail);
					if (iCount == 0) MessageBox.Show("El Proveedor no tiene Contactos con Direcciones de Correo registradas, \n\n NO SE ENVIARÁ EL CORREO DE ESTA ORDÉN DE COMPRA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					
				
					if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
					{									
						this.dtFechaLlegada.Value = System.DBNull.Value;
						this.dthorallegada.Value = System.DBNull.Value;
						this.txtIdComprobante.Value = 12;
						this.txtIdCredTributario.Value = 7;

						this.cmbImportacion.Enabled = true;

						this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = false;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = false;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescNC"].Hidden = true;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
						{
							dr.Cells["Impuesto"].Value = 0.00m;
							dr.Cells["DescNC"].Value = 0.00m;						
						}

						if (this.ultraGrid1.Rows.Count > 0) this.cmdTotal();

						if (miAcceso.CambiarFecha) this.dtFecha.Focus();
						else this.cmbImportacion.Focus();
					}
					else
					{
						this.cmbImportacion.Value = System.DBNull.Value;					
						this.txtIdComprobante.Value = 1;
						this.txtIdCredTributario.Value = 1;
						
						this.chkConsignacion.Enabled = true;					

						this.dthorallegada.Enabled = true;
											
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = true;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = true;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescNC"].Hidden = false;
						this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impuesto"].CellActivation = Activation.AllowEdit;

						if (iOrigen != 1)
						{
							this.dtFechaLlegada.Value = System.DBNull.Value;
							this.dthorallegada.Value = System.DBNull.Value;
							this.dtFechaLlegada.Enabled = true;
						}
						else
						{														
							this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);
							this.dthorallegada.CalendarInfo.MinDate = DateTime.Today.AddDays(2);
							DateTime fHoy1 = (DateTime) this.dthorallegada.Value;
							this.dthorallegada.CalendarInfo.MaxDate = fHoy1.Date.AddDays(2);
							//						this.dthorallegada.CalendarInfo.MinDate= DateTime.Today.AddDays(2);
						}
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
						{						
							dr.Cells["CantAereo"].Value = 0.00m;
							dr.Cells["Valores"].Value = 0.00m;
						}

						if (miAcceso.CambiarFecha) this.dtFecha.Focus();
						else this.dthorallegada.Focus();
					}
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) 
				{
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", this.txtRuc.Text.ToString())) == 0)
				{
					MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual Pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				
				if (this.chkConsignacion.Checked)
				{
					if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(AplicaConsignacion, 0) From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString()), true))
					{
						MessageBox.Show("El Proveedor no esta Asignado para Consignación. Revise la Ficha del Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();
						e.Cancel = true;
						return;
					}						
				}				
			}
		}

		private void dtFechaLlegada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.cmbImportacion.Focus();
				else
				{
					this.ultraGrid1.Focus();
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				}
			}
		}

		private void dtFechaLlegada_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.dtFechaLlegada.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha de Llegada de la Mercaderia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaLlegada.Focus();
				e.Cancel = true;				
			}
		}

		private void cmbImportacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.Focus();
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void cmbImportacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{
					if (this.cmbImportacion.Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione la Importación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbImportacion.Focus();
						e.Cancel = true;			
					}
				}		
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (ultraGrid1.ActiveCell != null) 
				{					
					FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	

					#region Articulos
					if (e.KeyCode == Keys.F3)
					{		
						try
						{	
							if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO")
							{									
								int idDetCompra = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
								int idArticulo = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

								string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
								int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										
								if (!bNuevo && (iConteo > 0 && !miAcceso.BModArtSerializado))
								{
									MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
									return;
								}
									
								FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, (int)this.cmbBodega.Value, 2, 6, Convert.ToDateTime(this.dtFecha.Value), false);									
							}						
						}
						catch(System.Exception ex)
						{
							MessageBox.Show(ex.Message, "Error en Ingreso de Productos", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					#endregion Articulos
				}
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					if (!bNuevo)
					{	
						for (int i=0; i<e.Rows.Length; i++)
						{							
							int idDetCompra = 0;							
							int idArticulo = 0;

							if (e.Rows[i].Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Rows[i].Cells["idDetCompra"].Value;
							if (e.Rows[i].Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Rows[i].Cells["idArticulo"].Value;

							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
							if (iConteo > 0)
							{
								if (!miAcceso.BEliArtSerializado)
								{
									MessageBox.Show(string.Format("Este Articulo ya tiene {0} Seriales Registradas no puede Eliminarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									e.Cancel = true;
								}
								else
								{
									if (DialogResult.No ==	MessageBox.Show(string.Format("¿Esta Seguro de Borrar el Artículo '{0}', con {1} Seriales, \n\n Este proceso no puede ser revertido?", e.Rows[i].Cells["Codigo"].Value.ToString(), iConteo), "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
									{
										e.Cancel = true;										
									}
									else
									{
										e.DisplayPromptMsg = false;

										string sSQLElimina = string.Format("Exec EliminaArticuloODCSeriales {0}, {1}, 2", (int)e.Rows[i].Cells["idDetCompra"].Value, iConteo);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);
									}
								}
							}
							else 
							{
								e.DisplayPromptMsg = false;

								string sSQLElimina = string.Format("Exec EliminaArticuloODCSeriales {0}, {1}, 2", (int)e.Rows[i].Cells["idDetCompra"].Value, iConteo);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);
							}
						}
					}
					else e.DisplayPromptMsg = false;
				}
				else e.DisplayPromptMsg = false;
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 2, (int)this.cmbBodega.Value);
				
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{
					e.Row.Cells["idImportacion"].Value = 1;
					e.Row.Cells["idTipoCalculo"].Value = Funcion.iEscalarSQL(cdsSeteoF, "Select idTipoCalculo From ImportacionTipo Where idImportacionTipo = 1");
				}
				else
				{
					e.Row.Cells["idImportacion"].Value = 0;
					e.Row.Cells["idTipoCalculo"].Value = 0;
				}
			}
		}

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

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value && dr.Cells["DescNC"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dr.Cells["Subtotal"].Value = Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescNC"].Value);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Subtotal"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2);
					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{						
						this.txtIva.Value = Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
						this.txtDescIvaTotal.Value = Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
					}
					else
					{				
						this.txtIva0.Value = Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
						this.txtDescIva0Total.Value = Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
					}
				}
								
				decimal Subtotal = Convert.ToDecimal(this.txtIva.Value);
				decimal Descuento = Convert.ToDecimal(this.txtDescIvaTotal.Value);
				decimal IVA = Math.Round((Subtotal - Descuento) * dPorcIva /100, 2);

				this.txtIvaTotal.Value = IVA;

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

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
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
			
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Valores", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Courier", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Peso", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Arancel", 5);
			FuncionesProcedimientos.FormatoGrid(e, "DescNC", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Subtotal", 5);			
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)
			{
				if (bActualiza)
				{
					bActualiza = false;
					#region Verifica Cambio
					int VerificaCambio = 0;
					#region Codigo o Articulo
					if (e.Cell.Column.ToString() == "Codigo")// || e.Cell.Column.ToString() == "Articulo")
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
							if (e.Cell.Row.Cells["Codigo"].Value.ToString().Length == 0) e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue.ToString();

							stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();						
						}
						#endregion Codigo
						#region Articulo
						if (e.Cell.Column.ToString() == "Articulo")
						{
							stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
						}
						#endregion Articulo

						int iTipoArticulo = 0;
								
						#region Cuenta coincidencias del codigo
						string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, {3}, 0", stCod, 2, (int) this.cmbBodega.Value, iTipoArticulo);
						if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
						iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
						#endregion Cuenta coincidencias del codigo
						#region Mas de una coincidencia o ninguna muestra la lista
						if (iCuenta != 1)
						{ 
							if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));
					
							int Bodega = (int) this.cmbBodega.Value;

							using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, Bodega, 2))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									if (bRastreo) Rastreo("Articulo escogido de la lista");
									IdArt = miBusqueda.idArticulo;
								}
								else
								{
									e.Cell.Row.Cells["idArticulo"].Value = 0;
									e.Cell.Row.Cells["Precio"].Value = 0.00m;
									e.Cell.Row.Cells["Codigo"].Value = "";
									e.Cell.Row.Cells["Articulo"].Value = "";
									cmdTotal();

									if (bRastreo) Rastreo("Cancela la selección de Artículos en lista");
									bActualiza = true;											
								}
							}
						}
							#endregion Mas de una coincidencia o ninguna muestra la lista
							#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
						else 
						{
							if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
							stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, (int) this.cmbBodega.Value, 2);
							IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
						}
						#endregion Encuentra un solo articulo
						#region Asignaciones
						if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
						stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 2, '{2}', 1", IdArt, (int)this.txtIdGrupoCliente.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						//@idArticulo Int, @idClienteGrupo Int, @idTipoFactura Int, @Fecha Datetime, @Fiscal Bit
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
						dr.Read();
						stCod = dr.GetString(0);
						e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
						e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
						e.Cell.Row.Cells["idArticulo"].Value = IdArt;
						e.Cell.Row.Cells["Precio"].Value = dr.GetDecimal(3);
						e.Cell.Row.Cells["Impuesto"].Value = dr.GetDecimal(4);						
						e.Cell.Row.Cells["Peso"].Value = dr.GetDecimal(5);						
						e.Cell.Row.Cells["Arancel"].Value = dr.GetDecimal(6);
						dr.Close();
						e.Cell.Row.Cells["Codigo"].Value = stCod;

						if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
						{
							string sSQL = string.Format("Exec ValInfoArticulosImportacion {0}, {1}", IdArt, (int)this.cmbImportacion.Value);
							string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
							if (sMensaje.Trim().Length > 0)
							{
								MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							}
						}
						VerificaCambio = 2; // Verificaciones y Cambio de precio
						#endregion Asignaciones
					} 
					#endregion Codigo o articulo
						
					#region Cantidad
					if (e.Cell.Column.ToString() == "Cantidad")
					{
						if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((int)e.Cell.Row.Cells["Cantidad"].Value > 1)
							{
								e.Cell.Value = 0;
								MessageBox.Show("Seleccione un Producto antes de Ingresar Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}

						e.Cell.Row.Cells["DescNC"].Value = 0.00m;

						#region Valida Seriales ingresadas
						if (!bNuevo)
						{
							int idDetCompra = 0;
							if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
							int idArticulo = 0;
							if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

							int iCantidadM = (int)e.Cell.Row.Cells["Cantidad"].Value;
 
							if (iCantidadM < iConteo)
							{
								e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
								MessageBox.Show(string.Format("El Artículo '{0}' tiene {1} seriales registradas", e.Cell.Row.Cells["Codigo"].Value.ToString().Trim(), iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
							}								
						}
						#endregion Valida Seriales ingresadas
						this.cmdTotal();
						VerificaCambio = 1;
					}
					#endregion Cantidad
					#region Cantidad Aereo
					if (e.Cell.Column.ToString() == "CantAereo")
					{
						if (e.Cell.Row.Cells["CantAereo"].Value == System.DBNull.Value) e.Cell.Row.Cells["CantAereo"].Value = (int)e.Cell.OriginalValue;

						string sSQLCantCmp = string.Format("Select dc.Cantidad From DetCompra dc Inner Join Compra c On c.idCompra = dc.idCompra Where dc.idOrigen = {0} And c.idTipoFactura = 4", (int)e.Cell.Row.Cells["idDetCompra"].Value);
						int dCantCMP = Funcion.iEscalarSQL(cdsSeteoF, sSQLCantCmp);
						int dCantAer = 0;
						if (e.Cell.Row.Cells["CantAereo"].Value != System.DBNull.Value) dCantAer = (int)e.Cell.Row.Cells["CantAereo"].Value;

						if ((dCantCMP - dCantAer) < 0)
						{
							e.Cell.Row.Cells["CantAereo"].Value = 0;
							MessageBox.Show(string.Format("La Cantidad Máxima es {0}", dCantCMP), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						else e.Cell.Row.Cells["Cantidad"].Value = dCantCMP - dCantAer;					
					}
					#endregion Cantidad Aereo								
					#endregion Fin de Verificacion de Cambio
					#region DescuentoPorc o DescuentoArt
					if (e.Cell.Column.ToString() == "DescuentoPorc")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["DescuentoPorc"].Value == System.DBNull.Value) e.Cell.Row.Cells["DescuentoPorc"].Value = (decimal)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}

						if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 100.00m)
						{								
							e.Cell.Value = 0.00m;
							MessageBox.Show("El Porcentaje de Descuento Maximo es 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
												
						cmdTotal();
						bActualiza = true;
					}

					if (e.Cell.Column.ToString() == "DescuentoArt")
					{
						bActualiza = false;
						cmdTotal();
						bActualiza = true;
					}
					#endregion DescuentoPorc o DescuentoArt
					#region Precio
					if (e.Cell.Column.ToString() == "Precio")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value) e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Precio"].Value > 0.00m)
							{
								e.Cell.Row.Cells["DescNC"].Value = 0.00m;
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}		
						
						cmdTotal();
						bActualiza = true;
					}
					#endregion Precio					
					#region IVA
					if (e.Cell.Column.ToString() == "Impuesto")
					{
						bActualiza = false;		
						if (e.Cell.Row.Cells["Impuesto"].Value == System.DBNull.Value) e.Cell.Row.Cells["Impuesto"].Value = Convert.ToDecimal(e.Cell.OriginalValue);

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
							if ((decimal)e.Cell.Row.Cells["Impuesto"].Value != dPorcIva)
							{

								e.Cell.Value = (decimal)e.Cell.OriginalValue;;
								MessageBox.Show(string.Format("El Porcentaje de I.V.A. debe ser 0% o {0}%", dPorcIva), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}
						cmdTotal();
						bActualiza = true;
					}
					#endregion IVA
					#region Descuento Promocion
					if (e.Cell.Column.ToString() == "DescNC")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["DescNC"].Value == System.DBNull.Value) e.Cell.Row.Cells["DescNC"].Value = (decimal)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["DescNC"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Descuento Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}		
						
						cmdTotal();
						bActualiza = true;
					}
					#endregion Descuento Promocion
					#region Valores
					if (e.Cell.Column.ToString() == "Valores")
					{
						bActualiza = false;	
						if (e.Cell.Row.Cells["Valores"].Value == System.DBNull.Value) e.Cell.Row.Cells["Valores"].Value = (decimal)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Valores"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar Valores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}
						bActualiza = true;
					}
					#endregion Valores
				}
				bActualiza = true;
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			this.ConsultaMaestro((int)this.txtIdCompra.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3 && !miAcceso.BAnularProcesado)
			{
				MessageBox.Show("No puede Anular una Orden de Compra PROCESADA \n\n Consulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if ((int)this.txtEstado.Value == 9 || (int)this.txtEstado.Value == 13 || (int)this.txtEstado.Value == 12)
			{
				MessageBox.Show("No puede Anular una Orden de Compra FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.txtIdCompra.Value, cdsSeteoF)) return;
			
			if (Funcion.ValidaBloqueo((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, 2", (int)this.txtIdCompra.Value), true) > 0 && !miAcceso.BAnularConSeriales)
			{
				MessageBox.Show("No puede Anular Existen Seriales Ingresadas\n\nConsulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}			
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Anular Esta Orden de Compra '{0}'", this.lblEstado.Text.ToString()), "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				if (Facturacion.AnulacionCompraVenta ((int)this.txtIdCompra.Value, 2, cdsSeteoF))
					this.ConsultaMaestro((int)this.txtIdCompra.Value);							
		}

		private void btnOrdenDeCompraMixta_Click(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if ((int)this.txtEstado.Value != 12)
				{
					MessageBox.Show("Debe haber Facturado para crear una Importación Mixta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				}
				else if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro de Crear Una Orden de Compra Mixta?", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					int iCant = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if ((int)dr.Cells["CantAereo"].Value == 0) iCant ++;	
					}

					if (iCant == this.ultraGrid1.Rows.Count)
					{
						MessageBox.Show("Escriba la cantidad de la nueva Importación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
					}
					else
					{
						#region Crea nueva Ordén de Compra
						string sSQL = string.Format("Exec CreaCopiaOrdCompraImpMixta {0}", (int)this.txtIdCompra.Value);
						int idCompraN = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

						if (idCompraN < 0)
						{
							MessageBox.Show("Ocurrió un Error al Grabar la Nueva Ordén de Compra.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
						else
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetCompra Where Cantidad = 0 And idCompra = {0}", (int)this.txtIdCompra.Value));
							
							Funcion.EjecutaSQL(cdsSeteoF, "Exec ImportacionNombreCarga");

							MessageBox.Show("Se creo correctamente la Orden de Compra.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);						
						}
						#endregion Crea nueva Ordén de Compra
					}
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{	
			this.ConsultaMaestro((int)this.txtIdCompra.Value);
			
			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3 && !miAcceso.BEditarProcesado)
			{
				MessageBox.Show("No puede editar una orden de compra PROCESADA, Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}

			if ((int)this.txtEstado.Value == 9 || (int)this.txtEstado.Value == 13 || (int)this.txtEstado.Value == 12)
			{
				MessageBox.Show("No Puede Editar una orden de compra FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			
			if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.txtIdCompra.Value, cdsSeteoF)) return;
			
			if (Funcion.ValidaBloqueo((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;
				
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			#region Controles
			this.txtRuc.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;			
			 
			if ((int)this.txtIdGrupoCliente.Value == 3 || (int)this.txtIdGrupoCliente.Value == 2)
			{				
				this.cmbImportacion.Enabled = true;
				if ((int)this.txtEstado.Value != 12) FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);				
			}
			else
			{
				FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

				this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);

				this.chkConsignacion.Enabled = true;
				//				this.dtFechaLlegada.Enabled = true;		
				this.dthorallegada.Enabled = true;
			}
						
			if ((int)this.txtEstado.Value == 1) this.cmbBodega.Enabled = true;
			
			this.txtNotas.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnOrdenDeCompraMixta.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnEMail.Enabled = false;

			bNuevo = false;
			bEdicion = true;
			#endregion Controles			
		}

		public void ConsultaMaestro(int IdCompra)
		{
			try
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;			
				this.btnOrdenDeCompraMixta.Enabled = false;
				this.btnEMail.Enabled = false;
				this.btnAnular.Enabled = false;				
				
				int IdBodegaC = 0;

				int idImportacion = 0;
				string sSQL = string.Format("Exec ConsultaMaestroOrdenDeCompra {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdCompra.Value = IdCompra;
					this.txtContadoCredito.Value = dr.GetInt32(1);
					this.txtIdComprobante.Value = dr.GetInt32(2); 
					this.txtIdCredTributario.Value = dr.GetInt32(3);
					this.chkConsignacion.Checked = dr.GetBoolean(4);
					this.txtIdCliente.Value = dr.GetInt32(5);
					this.txtIdTipoRuc.Value = dr.GetInt32(6);
					this.txtIdGrupoCliente.Value = dr.GetInt32(7);
					this.txtRuc.Text = dr.GetString(8);
					this.txtNombre.Text = dr.GetString(9);
					IdBodegaC = dr.GetInt32(10);
					this.dtFecha.Value = dr.GetDateTime(11);
					if (dr.GetValue(12) != System.DBNull.Value) this.dtFechaLlegada.Value = dr.GetDateTime(12);	
					//					if (dr.GetValue(13) != System.DBNull.Value)
					this.dthorallegada.Value = dr.GetDateTime(13);	
					this.lblUsuario.Text = dr.GetString(14);
					this.txtNumero.Text = dr.GetString(15);
					if (dr.GetValue(16) != System.DBNull.Value) idImportacion = dr.GetInt32(16);
					this.txtEstado.Value = dr.GetInt32(17);
					this.txtNotas.Text = dr.GetValue(18).ToString();
					this.txtIva0.Value = dr.GetDecimal(19);
					this.txtDescIva0Total.Value = dr.GetDecimal(20);
					this.txtIva.Value = dr.GetDecimal(21);
					this.txtDescIvaTotal.Value = dr.GetDecimal(22);
					this.txtIvaTotal.Value = dr.GetDecimal(23);
					this.txtTotal.Value = dr.GetDecimal(24);
					bAnulado = dr.GetBoolean(25);
				}
				dr.Close();
				this.cmbBodega.Value= IdBodegaC;

				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = false;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = false;
				}

				FuncionesProcedimientos.ListaImportacionesActivas(this.cmbImportacion, idImportacion, cdsSeteoF);
				this.cmbImportacion.Value = idImportacion;

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleOrdenCompra {0}", IdCompra));

				if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
				{
					if (bAnulado) this.lblEstado.Text = "ANULADO";
					else
					{
						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
						if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
						if ((int)this.txtEstado.Value == 9 || (int)this.txtEstado.Value == 13) this.lblEstado.Text = "FACTURADO";
						if ((int)this.txtEstado.Value == 12) this.lblEstado.Text = "FACTURADO SIN MERCADERIA";
					}

					if (!bAnulado)
					{
						if (miAcceso.Editar)
						{
							if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
							else if ((int)this.txtEstado.Value == 3 && miAcceso.BEditarProcesado) this.btnEditar.Enabled = true;
						}

						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

						if (miAcceso.Anular && !bAnulado)
						{
							if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;
							else if ((int)this.txtEstado.Value == 3 && miAcceso.BAnularProcesado)this.btnAnular.Enabled = true;
						}

						if ((int)this.txtEstado.Value == 1) this.btnEMail.Enabled = true;
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
			this.btnCancelar_Click(sender, e);
			
			using (frmBuscaOrdenCompra Busqueda = new frmBuscaOrdenCompra())
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.ConsultaMaestro((int)Busqueda.uGridOrdenesDeCompra.ActiveRow.Cells["idCompra"].Value);
		}

		private void btnEMail_Click(object sender, System.EventArgs e)
		{
			this.ConsultaMaestro((int)this.txtIdCompra.Value);
			//this.txtRuc.Enabled = true;

			string sSQLMail = string.Format("Select Count(*) From ClienteContacto Where EnvioMail = 1 And idCliente = {0}", (int)this.txtIdCliente.Value);
			int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLMail);

			if (iCount == 0) 
			{
				MessageBox.Show("El proveedor " + this.txtNombre.Text.ToString() + " no tiene contactos configurados para enviar E-Mail.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
			}
			else if (DialogResult.Yes == MessageBox.Show("¿Desea enviar E-Mail al proveedor " + this.txtNombre.Text.ToString() + " con esta Orden de Compra.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				Imprimir(-1);
			}			
		}

		private void Imprimir(int iTipo)
		{			
			string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.Value.ToString();
						
			using (Reporte miReporte = new Reporte("OrdenCompra.Rpt", stFiltro))
			{			
				if (iTipo == 0) miReporte.ShowDialog(); // Pantalla
				else if (iTipo == -1) // Archivo PDF
				{					
					miReporte.ExportaPDF(@"\OrdenesDeCompra\OrdenCompra_" + this.txtNumero.Text.ToString() + ".pdf");					
					
					string Archivo = this.txtNumero.Text.ToString() + ".pdf";
					
					string sSQLMail = string.Format(@"Exec MailOrdenCompra 'OrdenCompra_{0}', {1}", Archivo, (int)this.txtIdCompra.Value);					
					string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLMail, true);

					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			this.ActualizaIVA();
			this.cmdTotal();
		
			#region Valida Estado 
			if (!bNuevo)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					int idDetCompra = 0;
					if (dr.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)dr.Cells["idDetCompra"].Value;
					int idArticulo = 0;
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)dr.Cells["idArticulo"].Value;

					string sSQLValConteoSerial = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
					int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQLValConteoSerial);

					decimal dCantidadM = (int)dr.Cells["Cantidad"].Value;
 
					if (dCantidadM > iConteo)
					{						
						this.txtEstado.Value = 1;
						dr.Cells["EstadoSerial"].Value = false;						
					}
					else if (dCantidadM == iConteo)	dr.Cells["EstadoSerial"].Value = true;					
				}

				int iContEstado = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (!(bool)dr.Cells["EstadoSerial"].Value) iContEstado ++;					
				}

				if (iContEstado == 0) this.txtEstado.Value = 3;									
				else if (iContEstado > 0) this.txtEstado.Value = 1;					
			}
			#endregion Valida Estado 
			
			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) return;
							
			if (this.chkConsignacion.Checked && !Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AplicaConsignacion From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString()), true))
			{
				MessageBox.Show("El proveedor no esta asignado para Consignación. Revise la ficha de Cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();						
				return;
			}

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Compra", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Bodega")) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFechaLlegada, "Ingrese la fecha de llegada de la mercaderia", false, "Compra", cdsSeteoF) && ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)) return;
		
			if (!Validacion.vbFechaEnDocumentos(this.dthorallegada, "Ingrese la hora de llegada de la mercaderia", false, "Compra", cdsSeteoF) && ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)) return;
	
			if (((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) && !Validacion.vbComboVacio(this.cmbImportacion, "Seleccione la Importación")) return;
						
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 2)) return;			

			#region Valida Informacion de importacion de articulos
			if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
			{								
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					string sSQLVAlInf = string.Format("Exec ValInfoArticulosImportacion {0}, {1}", (int)dr.Cells["idArticulo"].Value, (int)this.cmbImportacion.Value);
					string sMensajes = Funcion.sEscalarSQL(cdsSeteoF, sSQLVAlInf);
					if (sMensajes.Trim().Length > 0)
					{											
						MessageBox.Show(sMensajes, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}							
			#endregion Valida Informacion de importacion de articulos			
									
			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACIÓN DEL PROVEEDOR \n PRODUCTOS \n PRECIOS CANTIDADES DESCUENTOS \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
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
						
						#region Numeracion
						if (bNuevo)
						{							
							oCmdActualiza.CommandText = "Exec FacturaNumeroEstacion 2";
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 						 
						}
						#endregion Numeracion

						#region Maestro
						bool bFechaLlegada = false;
						bool hFechaLlegada = false;
						DateTime dtFechaLlegada = DateTime.Today;
						DateTime dthorallegada= DateTime.Today;

						if (this.dtFechaLlegada.Value != System.DBNull.Value)
						{
							dtFechaLlegada = (DateTime)this.dtFechaLlegada.Value;
							bFechaLlegada = true;
						}
						if((this.dthorallegada.Value != System.DBNull.Value)&& (int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
						{
							dthorallegada=(DateTime)this.dthorallegada.Value;
							hFechaLlegada = true;
						}
						//						else
						//						{
						//							dthorallegada=DateTime.Today;
						//						}

						int idImportacion = 0;
						if (this.cmbImportacion.Value != System.DBNull.Value) idImportacion = (int)this.cmbImportacion.Value;
												
						string sSQLMaestro = string.Format("Exec GrabaMaestroOrdenDeCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', {11}, {12}, '{13}', {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}", 
							(int)this.txtIdCompra.Value, (int)this.txtContadoCredito.Value, (int)this.txtIdComprobante.Value, (int)this.txtIdCredTributario.Value, this.chkConsignacion.Checked,
							(int)this.txtIdCliente.Value, (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), dtFechaLlegada.ToString("yyyyMMdd HH:mm"), dthorallegada.ToString("yyyyMMdd HH:mm"), 
							idImportacion, (int)this.txtEstado.Value, this.txtNotas.Text.ToString(),
							(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
							bFechaLlegada, hFechaLlegada);
						oCmdActualiza.CommandText = sSQLMaestro;
						this.txtIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();						
						#endregion Maestro

						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{														
							string sSQLDetalle = string.Format("Exec GuardaDetalleOrdenDeCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}",
								(int)dr.Cells["idDetCompra"].Value, (int)this.txtIdCompra.Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, 														
								(int)dr.Cells["Cantidad"].Value, (int)dr.Cells["CantAereo"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, 
								(decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
								(int)this.txtIdProyecto.Value, (int)dr.Cells["idImportacion"].Value, 
								(int)this.cmbBodega.Value, 0, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["EstadoSerial"].Value,
								(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, 
								(decimal)dr.Cells["DescNC"].Value, (decimal)dr.Cells["Valores"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}

						if (MenuLatinium.bConsignacionVenta)
						{
							if (bNuevo)
							{
								oCmdActualiza.CommandText = "Exec FacturaNumeroEstacion 36";;
								string sNumeroNI = oCmdActualiza.ExecuteScalar().ToString(); 
							}
							
							int idCompraNI = 0;

							if (!bNuevo)
							{
								oCmdActualiza.CommandText = string.Format("If Exists(Select idCompra From Compra Where idTipoFactura = 36 And idOrigen = {0}) Select idCompra From Compra Where idTipoFactura = 36 And idOrigen = {0} Else Select CONVERT(Int, 0)", (int)this.txtIdCompra.Value);
								idCompraNI = (int)oCmdActualiza.ExecuteScalar();

								oCmdActualiza.CommandText = string.Format("Delete From DetCompra Where idCompra = {0}", idCompraNI); 
								oCmdActualiza.ExecuteNonQuery();
							}

							sSQLMaestro = string.Format("Exec GrabaMaestroNotaDeIngreso {0}, {1}, {2}, {3}, '{4}', '{5}', {6}, '{7}', {8}, {9}", 
								idCompraNI, this.chkConsignacion.Checked, (int)this.txtIdCliente.Value, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), 
								(int)this.txtEstado.Value, this.txtNotas.Text.ToString(), Convert.ToDecimal(this.txtTotal.Value), (int)this.txtIdCompra.Value);
							oCmdActualiza.CommandText = sSQLMaestro;
							idCompraNI = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{	
								string sSQLDetalle = string.Format("Exec GuardaDetalleNotaDeIngreso {0}, {1}, {2}, {3}, {4}, {5}",
									(int)dr.Cells["idDetCompra"].Value, idCompraNI, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Subtotal"].Value);
								oCmdActualiza.CommandText = sSQLDetalle;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
												 
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleOrdenCompra {0}", (int)this.txtIdCompra.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);
						this.ultraGrid1.DataSource = oDTDetalle;											
						#endregion Detalle	

						#region Turno Orden De Compra
						if (iOrigen == 1)
						{
							DateTime h30=dtFechaLlegada;

							while (h30 < (DateTime)this.dthorallegada.Value)
							{
								int iHoras = h30.Hour;
								int iMinutos = h30.Minute;
								string sCol = "T";
								string sHora = "";
								string sMin = "";
								if (iHoras.ToString().Length == 1) sHora = "0" + iHoras.ToString();
								if (iHoras.ToString().Length == 2) sHora = iHoras.ToString();
								if (iMinutos.ToString().Length == 1) sMin = iMinutos.ToString() + "0";
								if (iMinutos.ToString().Length == 2) sMin = iMinutos.ToString();
								sCol = sCol + sHora + sMin;
								string sSQLTurnoOC = string.Format("Exec ActualizaTurnoOrden '{0}', '{1}', {2}, {3}", sCol, dtFechaLlegada.ToString("yyyyMMdd HH:mm"), "1" + this.txtIdCompra.Value.ToString(), 1);
								h30= h30.AddMinutes(30);
								oCmdActualiza.CommandText = sSQLTurnoOC;
								oCmdActualiza.ExecuteNonQuery();
								string sSQLDetalleTurno = string.Format("Insert Into DetalleTurnoOrdenes (Fecha, valor,origen,columna) Values ('{0}', {1}, {2},'{3}')",dtFechaLlegada.ToString("yyyyMMdd HH:mm"), (int)this.txtIdCompra.Value, 1, sCol);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleTurno);
							}
						}
						#endregion Turno Orden De Compra

						#region Actualiza idArticulo en Seriales
						if (!bNuevo && miAcceso.BModArtSerializado)
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drAS in ultraGrid1.Rows.All)
							{
								string sSQLAS = string.Format("Exec ActualizaIdArticuloSerialesODC {0}", (int)drAS.Cells["idDetCompra"].Value);
								oCmdActualiza.CommandText = sSQLAS;
								oCmdActualiza.ExecuteNonQuery();
							}
						}												
						#endregion Actualiza idArticulo en Seriales											

						if (!bNuevo)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						oTransaction.Commit();

						if ((int)this.txtEstado.Value == 1) this.btnEMail_Click(sender, e);

						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
						if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
												
						#region Controles
						bNuevo = false;
						bEdicion = false;
						bEliminaAlValidar = false;

						this.chkConsignacion.Enabled = false;
						this.txtRuc.Enabled = false;
						this.dtFecha.Enabled = false;
						this.cmbBodega.Enabled = false;
						this.dtFechaLlegada.Enabled = false;
						this.dthorallegada.Enabled = false;								
						this.cmbImportacion.Enabled = false;
														
						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
																	
						this.txtNotas.Enabled = false;
						
						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						this.btnEMail.Enabled = true;
						this.btnGuardar.Enabled = false;												
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;																																																
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
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}

		private void btnImportacion_Click(object sender, System.EventArgs e)
		{
			using (ImportacionNombre miImportacionNombre = new ImportacionNombre())
			{
				if (DialogResult.OK == miImportacionNombre.ShowDialog())
				{
					Funcion.EjecutaSQL(cdsSeteoF, "Exec ImportacionNombreCarga");

					int idImportacion = 0;
					if (this.cmbImportacion.ActiveRow != null) idImportacion = (int)this.cmbImportacion.Value;
					FuncionesProcedimientos.ListaImportacionesActivas(this.cmbImportacion, idImportacion, cdsSeteoF);

					//this.cmbImportacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF,  "Select idImportacionNombre, Nombre From ImportacionNombre Order By idImportacionNombre");					
				}
			}
		}

		private void dtFechaLlegada_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void frmOrdenesDeCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
				if (e.KeyCode == Keys.X) if (miAcceso.Anular && bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
				if (e.KeyCode == Keys.M) if (this.btnEMail.Enabled) this.btnEMail_Click(sender, e);
				if (e.KeyCode == Keys.F5) 
				{
					miAcceso = new Acceso(cdsSeteoF, "0602");
					this.ConsultaMaestro((int)this.txtIdCompra.Value);
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

		private void chkConsignacion_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)
				this.dthorallegada.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{				
				if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
				{
					this.dtFechaLlegada.CalendarInfo.MinDate = (DateTime) this.dtFecha.Value;
					DateTime fHoy = (DateTime) this.dtFecha.Value;
					this.dtFechaLlegada.CalendarInfo.MaxDate = fHoy.Date.AddDays(90);
					//					this.dtFechaLlegada.Value = (DateTime) this.dtFecha.Value;


				}

				dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				ActualizaIVA();
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha del Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Ordenes de Compra en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;					
			}
		}

		private void cmbImportacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbImportacion_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbImportacion.Value != System.DBNull.Value) 
			{
				bool bCourier = false;
				
				int idImportacion = 0;
				
				if (this.cmbImportacion.ActiveRow != null) idImportacion = (int)this.cmbImportacion.Value;

				if (idImportacion > 0) bCourier = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Courier From ImportacionNombre Where idImportacionNombre = {0}", idImportacion), true);

				if (bCourier) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Courier"].Hidden = false;
				else this.ultraGrid1.DisplayLayout.Bands[0].Columns["Courier"].Hidden = true;
			}
			else this.ultraGrid1.DisplayLayout.Bands[0].Columns["Courier"].Hidden = true;
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{	
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

				string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
				int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										
				if (!bNuevo && (iConteo > 0 && !miAcceso.BModArtSerializado))
				{
					MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
					e.Cancel = true;
				}
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				if (iOrigen == 1)
				{
					this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BodegaUsuario 7");
					this.cmbBodega.Enabled = false;
					
				}
				else
				{
				
					FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 2, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
				}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) this.cmbImportacion.Focus();
				else this.dthorallegada.Focus();
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{			
			if (this.cmbBodega.ActiveRow != null) 
				this.txtIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
		}

		private void dthorallegada_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if ((int)this.txtIdGrupoCliente.Value != 2 && (int)this.txtIdGrupoCliente.Value != 3)
			{
				if((DateTime)this.dthorallegada.Value >=(DateTime)this.dtFechaLlegada.Value)
				{
					MessageBox.Show(string.Format("La hora de llegada debe ser mayor a la hore de inicio \n\nConsulte con el Administrador"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
				}
							
			}
		}				
	}
}

