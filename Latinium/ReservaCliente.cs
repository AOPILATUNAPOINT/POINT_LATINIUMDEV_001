using System;
using System.Drawing;
using System.Collections;
using System.Globalization;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Summary description for ReservaCliente.
	/// </summary>
	public class ReservaCliente : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtpHoraEntrega;
		private System.Windows.Forms.Label label67;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDirecciones;
		private System.Windows.Forms.Label label66;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtVence;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		[DllImport("user32.dll")]
		static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

		#region Variables Globales
		double MaximoCredito = 0;
		bool bEliminaAlValidar = false;
		private C1.Data.C1DataRow drCompra;
		private int IdSigno = 2;
		int IdTipoFactura = 11;
		//private C1.Data.C1DataSet cdsSeteo;
		private C1DataRow drSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label lblCodigoBodega;

		int IdBusqueda = 0;
		private bool bIncluyeIva = false;
		private C1.Data.C1DataSet cdsSeteoF;
		CultureInfo us = new CultureInfo("en-US");
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private bool bActualiza = true;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsCompraTabla;
		private Acceso miAcceso;
		bool bAutoImpresor = false;
		bool bNumeraImprimir = false;
		#endregion Variables Globales

		public ReservaCliente()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IdTipoFactura = 11;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// 

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetCompra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPeso");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaTelco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Elaboracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residuo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescGlobal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Razonable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CliFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoCDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PresupuestoC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PresupuestoCDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPasar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IceL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("costoBinaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("costoB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bordado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estampado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedioNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_IdArticulo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sugerido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Manual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantAereo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantMaritimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Liquidado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoBarras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("unidadBotrosa", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetRubro", 2);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCodigoBodega = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.dtpHoraEntrega = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label67 = new System.Windows.Forms.Label();
			this.cmbDirecciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label66 = new System.Windows.Forms.Label();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.dtVence = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbTipoRuc
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance1;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idTipoRuc"));
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(88, 6);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(112, 21);
			this.cmbTipoRuc.TabIndex = 196;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.Leave += new System.EventHandler(this.cmbTipoRuc_Leave);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCompra_PositionChanged);
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "_Compra";
			this.barraDato1.DataNombreId = "idCompra";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCompra;
			this.barraDato1.DataTabla = "Compra";
			this.barraDato1.DataTablaHija = "DetCompra";
			this.barraDato1.Location = new System.Drawing.Point(8, 504);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 195;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(800, 400);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(66, 16);
			this.lblSubTotal.TabIndex = 173;
			this.lblSubTotal.Text = "SUBTOTAL";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(416, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 172;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 170;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(800, 464);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(23, 16);
			this.lblIva.TabIndex = 168;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(800, 432);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(77, 16);
			this.lblDescuento.TabIndex = 166;
			this.lblDescuento.Text = "DESCUENTO";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(824, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 171;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance2;
			this.txtRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Ruc"));
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(208, 5);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 159;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(488, 5);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(312, 22);
			this.txtNombre.TabIndex = 158;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance4;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Total"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(904, 496);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(110, 34);
			this.txtTotal.TabIndex = 187;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance5;
			this.txtIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Iva"));
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(904, 464);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(110, 22);
			this.txtIvaTotal.TabIndex = 186;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance6;
			this.txtDescIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento"));
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(904, 432);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(110, 22);
			this.txtDescIvaTotal.TabIndex = 181;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance7;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalIva"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(904, 400);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(110, 22);
			this.txtIva.TabIndex = 179;
			// 
			// dtFecha
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance8;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(88, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 160;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Compra.Compra - DetCompra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance9;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 26;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 66;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 53;
			ultraGridColumn5.Header.Caption = "Artículo";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 194;
			ultraGridColumn6.Header.Caption = "Unidad";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn6.Width = 53;
			ultraGridColumn7.Header.VisiblePosition = 32;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 29;
			ultraGridColumn8.Hidden = true;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance10;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn9.NullText = "0";
			ultraGridColumn9.PromptChar = ' ';
			ultraGridColumn9.Width = 69;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance11;
			ultraGridColumn10.Format = "";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 75;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance12;
			ultraGridColumn11.Header.Caption = "Iva";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 51;
			ultraGridColumn12.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn12.Header.VisiblePosition = 63;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 24;
			ultraGridColumn13.Header.Caption = "Ref. Codigo";
			ultraGridColumn13.Header.VisiblePosition = 20;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 20;
			ultraGridColumn14.Header.Caption = "Ref. Numero";
			ultraGridColumn14.Header.VisiblePosition = 21;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 20;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Header.Caption = "Descuento";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 14;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance14;
			ultraGridColumn16.Header.Caption = "Peso";
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 17;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance15;
			ultraGridColumn17.Header.VisiblePosition = 28;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn17.PromptChar = ' ';
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance16;
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.MaskInput = "nnnnnnnnn";
			ultraGridColumn18.Width = 14;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance17;
			ultraGridColumn19.Header.Caption = "% Desc.";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 91;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance18;
			ultraGridColumn20.Header.VisiblePosition = 25;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn20.PromptChar = ' ';
			ultraGridColumn21.Header.VisiblePosition = 27;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 38;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 30;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 31;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 24;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.Caption = "Centro Costo";
			ultraGridColumn27.Header.VisiblePosition = 33;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn28.Header.Caption = "Proyecto";
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance19;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn29.NullText = "0";
			ultraGridColumn29.PromptChar = ' ';
			ultraGridColumn29.UseEditorMaskSettings = true;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 281;
			ultraGridColumn31.Header.VisiblePosition = 34;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.PromptChar = ' ';
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance20;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 14;
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.Width = 86;
			ultraGridColumn33.Header.Caption = "Importacion";
			ultraGridColumn33.Header.VisiblePosition = 35;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn34.Header.VisiblePosition = 36;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Width = 336;
			ultraGridColumn36.Header.Caption = "CentroCosto";
			ultraGridColumn36.Header.VisiblePosition = 13;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.Caption = "Proyecto";
			ultraGridColumn37.Header.VisiblePosition = 16;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 23;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 37;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.PromptChar = ' ';
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.PromptChar = ' ';
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.PromptChar = ' ';
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.Caption = "Vendedor";
			ultraGridColumn50.Header.VisiblePosition = 50;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn51.Header.Caption = "Obraid";
			ultraGridColumn51.Header.VisiblePosition = 49;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn52.Header.Caption = "Capituloid";
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn53.Header.Caption = "Rubroid";
			ultraGridColumn53.Header.VisiblePosition = 53;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn54.Header.Caption = "Obra";
			ultraGridColumn54.Header.VisiblePosition = 52;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.MaskInput = "";
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.MaskInput = "";
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn56.Header.Caption = "Rubro";
			ultraGridColumn56.Header.VisiblePosition = 57;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.MaskInput = "";
			ultraGridColumn57.Header.VisiblePosition = 61;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 56;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 59;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn60.Header.VisiblePosition = 55;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.MaskInput = "";
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 22;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.PromptChar = ' ';
			ultraGridColumn62.Width = 15;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 64;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 65;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 67;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 69;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 66;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 47;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 50;
			ultraGridColumn70.Header.VisiblePosition = 70;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 48;
			ultraGridColumn71.Header.VisiblePosition = 71;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 44;
			ultraGridColumn72.Header.VisiblePosition = 72;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 68;
			ultraGridColumn73.Header.VisiblePosition = 73;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 50;
			ultraGridColumn74.Header.VisiblePosition = 74;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 60;
			ultraGridColumn75.Header.VisiblePosition = 79;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 82;
			ultraGridColumn76.Header.VisiblePosition = 88;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 82;
			ultraGridColumn77.Header.VisiblePosition = 75;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 76;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 77;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 68;
			ultraGridColumn80.Header.VisiblePosition = 78;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 56;
			ultraGridColumn81.Header.VisiblePosition = 81;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 74;
			ultraGridColumn82.Header.VisiblePosition = 80;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 54;
			ultraGridColumn83.Header.VisiblePosition = 85;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 64;
			ultraGridColumn84.Header.VisiblePosition = 82;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 68;
			ultraGridColumn85.Header.VisiblePosition = 83;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 80;
			ultraGridColumn86.Header.VisiblePosition = 84;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 64;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 84;
			ultraGridColumn88.Header.VisiblePosition = 89;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 74;
			ultraGridColumn89.Header.VisiblePosition = 87;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 75;
			ultraGridColumn90.AutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			ultraGridColumn90.Header.Caption = "Código";
			ultraGridColumn90.Header.VisiblePosition = 3;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn90.Width = 214;
			ultraGridColumn91.Header.Caption = "Unidades";
			ultraGridColumn91.Header.VisiblePosition = 8;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn92.Header.VisiblePosition = 58;
			ultraGridColumn92.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
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
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 10F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance22;
			appearance23.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance23.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1010, 272);
			this.ultraGrid1.TabIndex = 161;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 16);
			this.label1.TabIndex = 197;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(800, 504);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 198;
			this.label2.Text = "TOTAL";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(824, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 200;
			this.label4.Text = "Número";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(416, 42);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 204;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance26;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idVendedor"));
			ultraGridColumn93.Header.VisiblePosition = 0;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 1;
			ultraGridColumn94.Width = 335;
			ultraGridColumn95.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(488, 40);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(312, 21);
			this.cmbVendedor.TabIndex = 203;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			this.cmbVendedor.Validated += new System.EventHandler(this.cmbVendedor_Validated);
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			this.cmbVendedor.Enter += new System.EventHandler(this.cmbVendedor_Enter);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdCompra.Appearance = appearance27;
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(320, 504);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 22);
			this.txtIdCompra.TabIndex = 207;
			this.txtIdCompra.Visible = false;
			// 
			// lblCodigoBodega
			// 
			this.lblCodigoBodega.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.lblCodigoBodega.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblCodigoBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblCodigoBodega.ForeColor = System.Drawing.Color.Black;
			this.lblCodigoBodega.Location = new System.Drawing.Point(912, 5);
			this.lblCodigoBodega.Name = "lblCodigoBodega";
			this.lblCodigoBodega.Size = new System.Drawing.Size(104, 22);
			this.lblCodigoBodega.TabIndex = 199;
			this.lblCodigoBodega.Text = "BI";
			this.lblCodigoBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			appearance28.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance28;
			this.txtNotas.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtNotas.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtNotas.Location = new System.Drawing.Point(8, 416);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(328, 80);
			this.txtNotas.TabIndex = 202;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// txtNumero
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance29;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(912, 39);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 209;
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
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// dtpHoraEntrega
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHoraEntrega.Appearance = appearance30;
			this.dtpHoraEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCompra, "_Compra.HoraEntrega"));
			this.dtpHoraEntrega.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtpHoraEntrega.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never;
			this.dtpHoraEntrega.FormatString = "HH:mm";
			this.dtpHoraEntrega.Location = new System.Drawing.Point(592, 432);
			this.dtpHoraEntrega.MaskInput = "{LOC}hh:mm";
			this.dtpHoraEntrega.Name = "dtpHoraEntrega";
			this.dtpHoraEntrega.Size = new System.Drawing.Size(80, 22);
			this.dtpHoraEntrega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtpHoraEntrega.TabIndex = 216;
			this.dtpHoraEntrega.Value = null;
			this.dtpHoraEntrega.Leave += new System.EventHandler(this.dtpHoraEntrega_Leave);
			this.dtpHoraEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpHoraEntrega_KeyPress);
			this.dtpHoraEntrega.Validating += new System.ComponentModel.CancelEventHandler(this.dtpHoraEntrega_Validating);
			this.dtpHoraEntrega.Validated += new System.EventHandler(this.dtpHoraEntrega_Validated);
			this.dtpHoraEntrega.ValueChanged += new System.EventHandler(this.dtpHoraEntrega_ValueChanged);
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(544, 432);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 16);
			this.label67.TabIndex = 215;
			this.label67.Text = "Hora";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDirecciones
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDirecciones.Appearance = appearance31;
			this.cmbDirecciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDirecciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDirecciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idDireccion"));
			this.cmbDirecciones.DisplayMember = "Direccion";
			this.cmbDirecciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDirecciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDirecciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDirecciones.Location = new System.Drawing.Point(416, 464);
			this.cmbDirecciones.Name = "cmbDirecciones";
			this.cmbDirecciones.Size = new System.Drawing.Size(352, 21);
			this.cmbDirecciones.TabIndex = 214;
			this.cmbDirecciones.ValueMember = "idDreccion";
			this.cmbDirecciones.Leave += new System.EventHandler(this.cmbDirecciones_Leave);
			this.cmbDirecciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbDirecciones_KeyPress);
			this.cmbDirecciones.Validating += new System.ComponentModel.CancelEventHandler(this.cmbDirecciones_Validating);
			this.cmbDirecciones.Validated += new System.EventHandler(this.cmbDirecciones_Validated);
			this.cmbDirecciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDirecciones_InitializeLayout);
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Location = new System.Drawing.Point(352, 464);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(50, 16);
			this.label66.TabIndex = 213;
			this.label66.Text = "Dirección";
			// 
			// chkEntregado
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance32;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.CausesValidation = false;
			this.chkEntregado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Entregado"));
			this.chkEntregado.Location = new System.Drawing.Point(352, 392);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(120, 22);
			this.chkEntregado.TabIndex = 212;
			this.chkEntregado.Text = "Entrega a domicilio";
			this.chkEntregado.CheckedChanged += new System.EventHandler(this.chkEntregado_CheckedChanged);
			this.chkEntregado.Leave += new System.EventHandler(this.chkEntregado_Leave);
			// 
			// dtFechaEntrega
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance33;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaEntrega"));
			dateButton2.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton2);
			this.dtFechaEntrega.Format = "dd/MM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(416, 432);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(104, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 211;
			this.dtFechaEntrega.Value = new System.DateTime(2014, 1, 21, 0, 0, 0, 0);
			this.dtFechaEntrega.Leave += new System.EventHandler(this.dtFechaEntrega_Leave);
			this.dtFechaEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaEntrega_KeyPress);
			this.dtFechaEntrega.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaEntrega_Validating);
			this.dtFechaEntrega.Validated += new System.EventHandler(this.dtFechaEntrega_Validated);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(352, 432);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(34, 16);
			this.label19.TabIndex = 210;
			this.label19.Text = "Fecha";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtComprobante
			// 
			this.txtComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance34;
			this.txtComprobante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprobante"));
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(912, 72);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(104, 22);
			this.txtComprobante.TabIndex = 243;
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(824, 75);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 242;
			this.lblComprobante.Text = "Comprobante";
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCompra, "_Compra.Usuario"));
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(480, 74);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 23);
			this.lblUsuario.TabIndex = 245;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(8, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 244;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 392);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(76, 16);
			this.label3.TabIndex = 246;
			this.label3.Text = "Observaciones";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(224, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(38, 16);
			this.label8.TabIndex = 248;
			this.label8.Text = "Vence:";
			// 
			// dtVence
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtVence.Appearance = appearance35;
			this.dtVence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtVence.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaVencimiento"));
			dateButton3.Caption = "Today";
			this.dtVence.DateButtons.Add(dateButton3);
			this.dtVence.Format = "dd/MM/yyyy";
			this.dtVence.Location = new System.Drawing.Point(280, 40);
			this.dtVence.Name = "dtVence";
			this.dtVence.NonAutoSizeHeight = 23;
			this.dtVence.Size = new System.Drawing.Size(112, 21);
			this.dtVence.SpinButtonsVisible = true;
			this.dtVence.TabIndex = 247;
			this.dtVence.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtVence.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtVence_KeyPress);
			this.dtVence.Validating += new System.ComponentModel.CancelEventHandler(this.dtVence_Validating);
			// 
			// ReservaCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1026, 534);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label67);
			this.Controls.Add(this.label66);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtVence);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.dtpHoraEntrega);
			this.Controls.Add(this.cmbDirecciones);
			this.Controls.Add(this.chkEntregado);
			this.Controls.Add(this.dtFechaEntrega);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.lblCodigoBodega);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "ReservaCliente";
			this.Text = "Reserva de Cliente";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ReservaCliente_KeyDown);
			this.Load += new System.EventHandler(this.ReservaCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompra_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_Compra")
			{
				drCompra = e.Row;
				if ((int) e.Row["idTipoFactura"] < 1)	
					e.Row["idTipoFactura"] = IdTipoFactura;
			}
		}

		#region Funciones Iniciales
		private void UnloadMe()
		{
			this.Close();
		}
		private void Seguridades()
		{
			#region Codigo de Seguridad
			string stTipo = "0707";
			string stPermiso = "122";
			
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			#endregion Tipo de Factura
			if (!Funcion.Permiso(stPermiso,cdsSeteoF))
			{
				string stSelect = "Select Nombre From CompraNumero Where idTipoFactura = " + IdTipoFactura;
				stSelect = Funcion.sEscalarSQL(cdsSeteo, stSelect);
				MessageBox.Show("No puede ingresar a " + stSelect, "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			#region Cambiar Fecha
			if (!miAcceso.CambiarFecha) 
			{
				dtFecha.ReadOnly = true;
				//toolTip.SetToolTip(dtFecha, "Fecha desactivada en seguridades del usuario. Consulte al administrador del sistema");
			}
			#endregion Cambiar Fecha
			#region Cambiar Numeracion
			if (!miAcceso.CambiarNumeracion) 
			{
				//toolTip.SetToolTip(txtNumero, "Numeración desactivada en seguridades del usuario. Consulte al administrador del sistema");
				txtNumero.ReadOnly = true;
			}
			#endregion Cambiar Numeracion
			#region Cambiar Precio
			// ! = Not && = And || = Or 
			if (!miAcceso.CambiarPrecio) ultraGrid1.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			#endregion Cambio de Precio
			#region Cambiar Descuento
			if (!miAcceso.CambiarDescuento)
			{
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoArt"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			}
			#endregion Cambio de Descuento
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

		private void EstadoControles(bool bHabiltado)
		{
			foreach (Control X in Controls)
			{
				if (!(X is Infragistics.Win.Misc.UltraButton) && !(X is Label) && !(X is BarraDatoSQL.BarraDatoSQL) && !(X is GroupBox))X.Enabled = bHabiltado;
			}
		}

		private void Configuracion()
		{
			#region Asigna Signo
			string stSelect = "SELECT Signo FROM CompraNumero WHERE idTipoFactura = " + IdTipoFactura;
			IdSigno = Funcion.iEscalarSQL(cdsCompra, stSelect);
			#endregion Asigna Signo
      barraDato1.stFiltroExt = " And idTipoFactura = " + IdTipoFactura;
			#region Lectura de Archivo - Propiedades - Configuracion
//			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
//			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#endregion
			#region Incluye Iva - ICE
			if ((bool) drSeteoF["IncluyeIva"]) bIncluyeIva = true;
			MostrarIce = (bool) drSeteoF["IceLinea"];
			#endregion Incluye Iva - ICE 		
		}
		#endregion Funciones Iniciales

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void ReservaCliente_Load(object sender, System.EventArgs e)
		{
			Seguridades();
			Configuracion();

			#region Busqueda de Registro
			if (IdBusqueda > 0)
			{
				barraDato1.IdRegistro = IdBusqueda;
				barraDato1.ProximoId(5);
			}
			else
				barraDato1.PosUltima();
			#endregion Busqueda de Registro

			cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
			
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec PersonalVendedor 1");
			cmbVendedor.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;

			#region Habilita controles
			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);
			#endregion habilita controles
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			#region Seguridad
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}			
			#endregion Seguridad

			this.barraDato1.CrearRegistro();
			//EstadoControles(false);

			#region Predefinidos
			drCompra["idTipoRuc"] = 4;
			drCompra["Fecha"] = DateTime.Now;
			drCompra["Estado"] = 1;
			drCompra["idComprobante"] = 15;
			drCompra["idTipoFactura"] = IdTipoFactura;
			drCompra["idProyecto"] = Funcion.iEscalarSQL(cdsCompra, "Exec ProyectoPredefinido 11, 1", true);
			drCompra["Bodega"] = Funcion.iEscalarSQL(cdsCompra, "Exec ProyectoPredefinido 11, 3", true);	
			this.dtFechaEntrega.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			#endregion Predefinidos

			#region Controles
			this.txtNombre.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtComprobante.Enabled = false;			
			this.txtIva.Enabled = false;
			this.txtDescIvaTotal.Enabled = false;			
			this.txtIvaTotal.Enabled = false;
			this.txtTotal.Enabled = false;
			ModificaGrilla(true);
			if (!miAcceso.CambiarFecha) this.dtFecha.Enabled = false;
			if (!miAcceso.BCambiarVencimiento) this.dtVence.Enabled = false;
			else 
			{
				this.dtVence.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			}

			DateTime dtVence = (DateTime)drCompra["Fecha"];
			if (!miAcceso.CambiarFecha && !miAcceso.BCambiarVencimiento) drCompra["FechaVencimiento"] = dtVence.AddDays(8);
			if (!miAcceso.CambiarFecha && miAcceso.BCambiarVencimiento)
			{
				drCompra["FechaVencimiento"] = dtVence.AddDays(8);
				this.dtVence.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			}
			#endregion Controles

			#region Precio por Forma de Pago
			if ((int) drSeteoF["TipoPrecio"] == 2)
			{
				using (CompraFormaPagoPrecio miFormaPago = new CompraFormaPagoPrecio(IdTipoFactura))
				{
					if (DialogResult.OK == miFormaPago.ShowDialog())
					{							
						drCompra["idFormaPago"] = int.Parse(miFormaPago.IdFormaPago.Value.ToString());
							
						if (int.Parse(miFormaPago.IdFormaPago.Value.ToString()) == 6) drCompra["ContadoCredito"] = 2;
						else if (int.Parse(miFormaPago.IdFormaPago.Value.ToString()) != 6) drCompra["ContadoCredito"] = 1;
												
						this.txtRuc.Focus();
					}
					else
						this.barraDato1_Deshacer(sender, e);
				}
			}
			#endregion Precio por Forma de Pago

			this.txtRuc.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			try
			{
				drCompra = cdsCompra.TableViews["Compra"].Rows[0];

				this.lblEstado.Text = "";
				
				if ((bool)drCompra["Borrar"]) this.lblEstado.Text = "ANULADO";
				else 
				{
					if ((int)drCompra["Estado"] == 1) lblEstado.Text = "PENDIENTE";
					if ((int)drCompra["Estado"] == 2) lblEstado.Text = "EN PROCESO";
					if ((int)drCompra["Estado"] == 3) lblEstado.Text = "PROCESADO";
					if ((int)drCompra["Estado"] == 4) lblEstado.Text = "RECHAZADO";
					if ((int)drCompra["Estado"] == 5) lblEstado.Text = "APROBADO";
					if ((int)drCompra["Estado"] == 9) lblEstado.Text = "FACTURADO";
				}

				#region Direcciones
				this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", (int)drCompra["idCliente"]));
				this.cmbDirecciones.DisplayLayout.Bands[0].Columns["idDreccion"].Hidden = true;
				this.cmbDirecciones.DisplayLayout.Bands[0].Columns["Direccion"].Width = 360;
				#endregion Direcciones
			}
			catch{};
			barraDato1.ToolTipBuscar(Funcion.Tiempo("Buscar"));
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
			ModificaGrilla(false);
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drCompra == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
				
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede Editar esta Reserva, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;										
			}

			this.barraDato1.IdRegistro = (int)drCompra["idCompra"];
			this.barraDato1.ProximoId(5);

			if ((bool)drCompra["Borrar"])
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)drCompra["Estado"] == 3 && !miAcceso.BEditarProcesado)
			{
				MessageBox.Show("No puede Editar la Reserva ya esta PROCESADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)drCompra["Estado"] == 5 && !miAcceso.BEditarAprobado)
			{
				MessageBox.Show("No puede editar la Reserva ya esta APROBADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)drCompra["Estado"] == 9)
			{
				MessageBox.Show("No puede editar la Reserva ya esta FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}			

			this.barraDato1.EditarRegistro();			
			
			ModificaGrilla(true);

			#region Controles del encabezado deshabilitados
			this.txtNumero.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbVendedor.Enabled = false;
			if (!miAcceso.CambiarFecha) this.dtFecha.Enabled = false;
			if (!miAcceso.BCambiarVencimiento) this.dtVence.Enabled = false;
			#endregion Controles del encabezado deshabilitados
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			cmdTotal();
			#region Valida
			#region RUC
			if (drCompra["RUC"] == DBNull.Value)
			{
				MessageBox.Show(string.Format("Digite el número de '{0}'", this.cmbTipoRuc.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtRuc.Focus();
				return;
			}
			#endregion RUC
			#region Busca Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', 11, {1}", this.txtRuc.Text.ToString(), (int)drCompra["Bodega"]);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (!drC.HasRows)
			{
				drC.Close();
				MessageBox.Show("Cliente no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtRuc.Focus();
				return;
			}
			drC.Close();
			#endregion Busca Proveedor
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Reservas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;										
			}
			#region Vendedor
			if (drCompra["idVendedor"] == DBNull.Value || (int)drCompra["idVendedor"] == 0)
			{
				MessageBox.Show("Seleccione un vendedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbVendedor.Focus();
				return;
			}
			#endregion Vendedor
			#region Valida grid vacia
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los artículos para la reserva.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Valida grid vacia
			#region Valida Entrega
			if ((bool)this.chkEntregado.Checked)
			{
				if (drCompra["FechaEntrega"] == System.DBNull.Value)
				{
					MessageBox.Show("Escriba la fecha de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.dtFechaEntrega.Focus();
					return;
				}
				if (drCompra["HoraEntrega"] == System.DBNull.Value)
				{
					MessageBox.Show("Escriba la hora de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.dtpHoraEntrega.Focus();
					return;
				}
				if (this.cmbDirecciones.Text.ToString().Length == 0)
				{
					MessageBox.Show("Seleccione la dirección de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbDirecciones.Focus();
					return;
				}
			}
			#endregion Valida Entrega			
			#region Valida Cantidad y precio de los articulos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				#region Borra Filas que no tienen articulos ingresados
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value || dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
				#endregion Borra Filas que no tienen articulos ingresados

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int IdArt = (int) dr.Cells["idArticulo"].Value;
					
					#region Valida Cantidad que sea mayor a 0
					double dCantidad = (double) dr.Cells["cantidad"].Value;
					if (dCantidad == 0)
					{
						MessageBox.Show(string.Format("Digite la cantidad del artículo '{0}'.", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
					#endregion Valida Cantidad que sea mayor a 0

					#region Valida que el Precio sea mayor a 0
					double dPrecio = (double) dr.Cells["Precio"].Value;
					if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value) 
					{
						MessageBox.Show(string.Format("Digite el precio del artículo '{0}'.", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
					#endregion Valida que el Precio sea mayor a 0
				}
			}
			#endregion Valida Cantidad y precio de los articulos
			#region Verifica Saldo Cero
			if ((double) txtTotal.Value == 0) 
			{
				MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Verifica Saldo Cero			
			#endregion Valida
			#region Numeracion
			if (barraDato1.bNuevo) drCompra["Numero"] = Funcion.sEscalarSQL(cdsCompra, "EXEC FacturaNumeroEstacion " + IdTipoFactura);			
			#endregion Numeracion		

			if ((int)drCompra["idFormaPago"] == 6) drCompra["Estado"] = 1;
			else drCompra["Estado"] = 5;

			this.barraDato1.GrabaRegistro();

			this.barraDato1.IdRegistro = (int)drCompra["idCompra"];
			this.barraDato1.ProximoId(5);

			ModificaGrilla(false);
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{			
				if (miAcceso.CambiarFecha) this.dtFecha.Focus();	 
				else if (miAcceso.BCambiarVencimiento) this.dtVence.Focus();	 
				else this.cmbVendedor.Focus();
			}
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				int iBod = 0;
				if (drCompra["Bodega"] != System.DBNull.Value) iBod = (int) drCompra["Bodega"];

				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), IdTipoFactura, (int)drCompra["idFormaPago"], (int)drCompra["idTipoRuc"], iBod, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						drCompra["Ruc"] = miBusqueda.Ruc;
						drCompra["idTipoRuc"] = miBusqueda.idTipoRuc;

						if (miAcceso.CambiarFecha) this.dtFecha.Focus();	 
						else if (miAcceso.BCambiarVencimiento) this.dtVence.Focus();	 
						else this.cmbVendedor.Focus();
					}
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
						
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
			this.txtTotal.FormatString = stFormato;
			
			this.txtDescIvaTotal.MaskInput = stInput;
			
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			
			this.txtTotal.MaskInput = stInput;
			#endregion
			
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].Format = stFormato;
			e.Layout.Bands[0].Columns["CantidadPeso"].Format = stFormato;
			e.Layout.Bands[0].Columns["Pedido"].Format = stFormato;
			e.Layout.Bands[0].Columns["Unidades"].Format = stFormato;

			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["CantidadPeso"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Pedido"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Unidades"].MaskInput = stInput;
			#endregion
			
			#region Decimales para el precio
			stFormato = stMoneda + stCeros.Substring(0,  4); //(int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0,  4); //(int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;
			#endregion
			
			#region Decimales para el Impuesto
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;
			#endregion

			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, 4); //(int) drSeteoF["decPorcDesc"]);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, 4); // iDecPorD);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			#endregion Decimales para porcentaje de descuento
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{
				#region F3 Inventario
				case (int) Keys.F3:				
					if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && barraDato1.bEditar)
					{
						if (!this.barraDato1.bNuevo)
						{
							int idDetCompra = 0;
							if (this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
							int idArticulo = 0;
							if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
							if (iConteo > 0)
							{
								MessageBox.Show("Este Articulo ya tiene las Seriales Registradas no puede Modificarlo \n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}

						string sArticulo = ""; 
						sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						int IdArt = 0;
					
						using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, 1, 11))
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								IdArt = miBusqueda.idArticulo;
								ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
							}
						}		
					}
					break;
				#endregion F3 Inventario	
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
				if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
			//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
				break;
				#endregion Enter
			}
		}

		private void cmdTotal()
		{
			#region Creacion de Variables
			int IdArticulo = 0;
			double dSubtIva = 0;
			double dSubtIva0 = 0;
			double dSubTotal = 0;
			double dImpuesto = 0;
			double dDescIva = 0;
			double dDescIva0 = 0;
			double dDescuento = 0;
			double dCantidad = 0;
			double dPrecio = 0;
			double dIva = 0;
			double dIce = 0;
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				dIva = 0;
				dDescuento = 0;
				if (dr.Cells["idArticulo"].Value != DBNull.Value) IdArticulo = (int)dr.Cells["idArticulo"].Value;
				if (dr.Cells["Cantidad"].Value != DBNull.Value) dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != DBNull.Value)	dPrecio = (double) dr.Cells["Precio"].Value;
				if (dr.Cells["Impuesto"].Value != DBNull.Value)	dIva = (double) dr.Cells["Impuesto"].Value;
				if (dr.Cells["DescuentoArt"].Value != DBNull.Value)	dDescuento = (double) dr.Cells["DescuentoArt"].Value;
				if (dr.Cells["IceL"].Value != DBNull.Value) 
					dIce += (double) dr.Cells["IceL"].Value * (dCantidad * dPrecio - dDescuento) / 100.0;
				
				dSubTotal = dCantidad * dPrecio + .000000001;
				
				dSubTotal = Math.Round(dSubTotal, dRedondeo);

				if (dIva > 0)
				{
					if (bIncluyeIva)
					{
						double dIvaLinea = Math.Round(dSubTotal - dSubTotal /(1+dIva/100.0), dRedondeo);
						double dSubTotalL = dSubTotal - dIvaLinea;
						dSubtIva += dSubTotal - dIvaLinea;
						dImpuesto += (dSubTotalL - dDescuento) * dIva / 100.0;
					}
					else
					{
						dSubtIva += Math.Round(dSubTotal, dRedondeo);
						dImpuesto += (dSubTotal - dDescuento) * dIva / 100.0;
					}

					dDescIva += dDescuento;
				}
				else
				{
					dSubtIva0 += dSubTotal;
					dDescIva0 += dDescuento;
				}
			} 
			#endregion Desplazamiento de grilla
			#region Calculo de Descuentos

			drCompra["Desc3"] = dDescIva;
			drCompra["Desc4"] = dDescIva0;

			dDescuento += dDescIva0 + dDescIva; // descuento total de la factura

			dDescuento = dDescIva0 + dDescIva;
			dDescIva = (double) drCompra["Desc1"];
			dDescIva0 = (double) drCompra["Desc2"];
			drCompra["Descuento"] = dDescIva + (double) drCompra["Desc3"];
			drCompra["Descuento0"] = dDescIva0 + (double) drCompra["Desc4"];
			
			dDescuento = Math.Round(dDescuento, dRedondeo);
			#endregion Calculo de Descuentos
			#region Calculo de subtotales
			dSubtIva0 += 0;
			dSubtIva0 = Math.Round(dSubtIva0, dRedondeo);
			dSubtIva = Math.Round(dSubtIva, dRedondeo);
		
			drCompra["SubTotalExcento"] = dSubtIva0;
			drCompra["SubTotalIva"] = dSubtIva;
						
			if ((bool) drSeteoF["IncluyeIva"]) drCompra["SubTotalIva"] = dSubtIva;
			else drCompra["SubTotalIva"] = dSubtIva;

			#region no se que significa revisar posteriormente
			double dSubtIvaAntes = dSubtIva - (double) drCompra["Desc3"];
			if ((double) drCompra["Desc1"] > 0 && dSubtIvaAntes > 0)
				dImpuesto -= dImpuesto * (1 - (dSubtIvaAntes - (double) drCompra["Desc1"])/ dSubtIvaAntes);
			dImpuesto = Math.Round(dImpuesto, dRedondeo);
			#endregion no se que significa revisar posteriormente
			drCompra["Iva"] = dImpuesto;
			#endregion Calculo de subtotales
			#region Calculo del Total
			double dTotal = 0;
			
			dTotal = dSubtIva0 + dSubtIva + dImpuesto - dDescuento; 
				
			dTotal = Math.Round(dTotal, dRedondeo);

			if(dTotal < 0)	
			{
				drCompra["Total"] = 0;
				errorProvider.SetError(txtTotal, "Error en el calculo del total: " + dTotal.ToString("#,##0.00"));
			}
			else
			{
				drCompra["Total"] = dTotal;
				errorProvider.SetError(txtTotal, ""); 
			}
			#endregion Calculo del Total
		}

		private void CambiaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((bool) drSeteoF["VerUnidad"] == false) return;
			int IdUnidad = (int) e.Cell.Row.Cells["idUnidad"].Value;
			int IdArticulo = (int) e.Cell.Row.Cells["idArticulo"].Value;
			double dUnidades = 0;
			if (e.Cell.Row.Cells["Unidades"].Value != DBNull.Value)
				dUnidades = (double) e.Cell.Row.Cells["Unidades"].Value;
			else
				e.Cell.Row.Cells["Unidades"].Value = 0;
		
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVUNP'")) 				
			{
				string stExec = string.Format("Exec ArticuloUnidadCambio {0}, {1}", 
					IdArticulo, IdUnidad);
				e.Cell.Row.Cells["Cantidad"].Value = Funcion.dEscalarSQL(cdsCompra, stExec) * dUnidades;

				return;
			}
			string stSelect = "Exec UnidadCambio "	+ IdArticulo;
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsSeteoF, stSelect, false);
			miReader.Read();

			int IdUnidadOr = miReader.GetInt32(0);
			//Poner solo cuando hay unidades visible
			if (IdUnidad == IdUnidadOr)
			{
				double dCant = miReader.GetDouble(4);
				e.Cell.Row.Cells["Cantidad"].Value = 
					(double) e.Cell.Row.Cells["Unidades"].Value * dCant;
			}
			else
			{
				double dCantidad = miReader.GetDouble(1); // Unidad Alterna
				if (dCantidad > 0)
				{
					e.Cell.Row.Cells["Cantidad"].Value = 
						(double) e.Cell.Row.Cells["Unidades"].Value * dCantidad;
					bool bPeso = miReader.GetBoolean(5);
					if (bPeso) // Venta por peso cambia el precio
					{
						e.Cell.Row.Cells["Precio"].Value = 
							(double) e.Cell.Row.Cells["Precio"].Value / dCantidad;
						e.Cell.Row.Cells["Precio"].Refresh();
					}
				}
				if (dCantidad == 0)
				{
					e.Cell.Row.Cells["Cantidad"].Value = (double) e.Cell.Row.Cells["Unidades"].Value;
				}
			}
			miReader.Close();
		}

		private void Precio(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return;

			#region Variables
			int IdArt = (int) e.Cell.Row.Cells["idArticulo"].Value;
			int idCliente = 0;
			if (drCompra["idCliente"] != DBNull.Value) 
				idCliente = (int) drCompra["idCliente"];
			double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			int idFormaPago = (int) drCompra["idFormaPago"];
			#endregion Variables
			#region Lee Precio 
			string stExec = string.Format("Exec ArticuloPrecio {0}, {1}, {2}, {3}, {4}, {5}", IdArt, idCliente, dCantidad.ToString("0.000000", us), idFormaPago, IdTipoFactura, (int)drCompra["Bodega"]);
			double fPrecio = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			if (fPrecio >=0) e.Cell.Row.Cells["Precio"].Value = fPrecio;
			#endregion Lee Precio
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string st = "";
			if (e.Cell.Column.ToString() == "Codigo")
			{
				st = e.Cell.Row.Cells["Codigo"].Value.ToString();
			}

			if (!bActualiza) return;
			
			#region Cuando Cambia CodArticulo a idArticulo
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "CodArt")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idArticulo"].Value = e.Cell.Row.Cells["CodArt"].Value;
					}
			#endregion cuando Cambia CodArticulo a idArticulo
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
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {1}, 0, 0", stCod, IdTipoFactura, (int)drCompra["Bodega"]);
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				#region Mas de una coincidencia o ninguna muestra la lista
				if (iCuenta != 1)
				{ 
					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, (int)drCompra["Bodega"], 11))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdArt = miBusqueda.idArticulo;
						}
						else
						{
							e.Cell.Row.Cells["idArticulo"].Value = 0;
							e.Cell.Row.Cells["Precio"].Value = 0.0;
							e.Cell.Row.Cells["Codigo"].Value = "";
							e.Cell.Row.Cells["Articulo"].Value = "";
							cmdTotal();

							bActualiza = true;
							return;						
						}
					}
				}
					#endregion Mas de una coincidencia o ninguna muestra la lista
				#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
				else 
				{
					stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, (int)drCompra["Bodega"], IdTipoFactura);
					IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
				}
				#endregion Encuentra un solo articulo

				#region Asignaciones
				stSelect = "Select Codigo, Articulo, idUnidad From Articulo Where idArticulo =" + IdArt;
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
				dr.Read();
				e.Cell.Row.Cells["Codigo"].Value = dr.GetString(0);
				e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
				e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				dr.Close();

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
				VerificaCambio = 1; // Verificaciones y cambio de precio conficional
			}
			#endregion Cantidad
			#region Unidad
			if (e.Cell.Column.ToString() == "Unidad")
			{
				MessageBox.Show("Cambio en Unidad!");
				if ((int) drSeteoF["TipoPrecio"] == 1 || (int) drSeteoF["TipoPrecio"] == 3 || (int) drSeteoF["TipoPrecio"] == 2) Precio(e);
			}
			#endregion Unidad
			#region Verificaciones de cambio de articulo
			if (VerificaCambio > 0)
			{
				#region Precio
				if (VerificaCambio == 2) // Cambia de Precio
				{
					Precio(e);
				}
				else
				{
					// Cambio de precio por cantidad o por procedimiento
					if ((int) drSeteoF["TipoPrecio"] == 1 || (int) drSeteoF["TipoPrecio"] == 3) 
					{						
						Precio(e);
					}
					else					
					{
						Precio(e);
					}
				}
				#endregion Precio
				//DescuentoPorcentaje(e);
				//Promocion(e);
				#region Limites
				
					if (!Limites(e)) 
					{
						bActualiza = true;
						return;
					}
				
				#endregion Limites
				
				cmdTotal();
			}
			#endregion Verificaciones
					
			#endregion Fin de Verificacion de Cambio
			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (!this.barraDato1.bNuevo)
				{
					int idDetCompra = 0;
					if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
					int idArticulo = 0;
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

					string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
					int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					double dCantidadM = (double)e.Cell.Row.Cells["Cantidad"].Value;
 
					if (dCantidadM < iConteo)
					{
						e.Cell.Row.Cells["Cantidad"].Value = (double)e.Cell.OriginalValue;
						MessageBox.Show(string.Format("La Cantidad del Artículo '{0}' No Puede Ser Menor a {1}.", e.Cell.Row.Cells["Codigo"].Value.ToString().Trim(), iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
					}					
				}

				VerificaCambio = 1;
			}
			#endregion Cantidad
			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				bActualiza = false;
				//DescuentoPorcentaje(e);
				
				cmdTotal();
				bActualiza = true;
			}
			if (e.Cell.Column.ToString() == "DescuentoArt")
			{
				double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
				double dPrecio = 0;
				if (e.Cell.Row.Cells["Precio"].Value != DBNull.Value)
					dPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
				if (dCantidad == 0 || dPrecio == 0) 
				{
					bActualiza = true;
					return;
				}

				double dDesc = dCantidad * dPrecio * (double) 
					e.Cell.Row.Cells["DescuentoPorc"].Value / 100.0;
				double dDesc1 = Math.Round((double) e.Cell.Row.Cells["DescuentoArt"].Value, 2);
				double dPorc = dDesc1 / dCantidad / dPrecio *100.0;
				dPorc = Math.Round(dPorc, 2);
				if ((double) drSeteoF["LimiteDescuento"] >= 0.01 && dPorc > (double) drSeteoF["LimiteDescuento"])
				{
					MessageBox.Show("Descuento no Permitido", "Información");
					e.Cell.Row.Cells["DescuentoArt"].Value = 0;
					e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				}
				else
				{
					if (Math.Abs(dDesc - dDesc1) > .009)
						e.Cell.Row.Cells["DescuentoPorc"].Value = dPorc;
				}
				cmdTotal();
			}
			#endregion DescuentoPorc o DescuentoArt
			#region Precio
			if (e.Cell.Column.ToString() == "Precio")
			{
				DescuentoPorcentaje(e);
				cmdTotal();
			}
			#endregion Precio

			bActualiza = true;
		}

		private void DescuentoPorcentaje(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Variables
			if ((double) e.Cell.Row.Cells["DescuentoPorc"].Value < 0.00) return;
			double dDescuentoPorc = (double) e.Cell.Row.Cells["DescuentoPorc"].Value;
			double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			double dIva = (double) e.Cell.Row.Cells["Impuesto"].Value;
			double dPrecio = 0;
			if (e.Cell.Row.Cells["Precio"].Value != DBNull.Value) dPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
			#endregion Variables
			#region Mensajes
			string stExec = string.Format("Exec ArticuloDescPorc {0}, {1}, {2}, {3}, {4}", dDescuentoPorc.ToString("0.000000"), IdTipoFactura, dCantidad.ToString("0.000000", us), dPrecio.ToString("0.00000", us), dIva.ToString("0.00", us));
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);
			if (stMensaje.Length > 0)
			{
				e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				MessageBox.Show("Descuento Supera el Limite de Descuento Permitido", "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			
			if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ArticuloAplicaDescuento {0}, {1}", (int)e.Cell.Row.Cells["idArticulo"].Value, (int)drCompra["idFormaPago"]), true))
			{
				e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				MessageBox.Show("El artículo no aplica descuento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}			
			#endregion Mensajes
			#region Valor
			e.Cell.Row.Cells["DescuentoArt"].Value = Funcion.dEscalarSQL(cdsSeteoF, stExec + ", 1");
			#endregion 
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
			string stExec = string.Format("Exec ArticuloIva {0}, {1}, {2}", 	IdTipoFactura, stIdArt, 0);
			double fIva = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			e.Cell.Row.Cells["Impuesto"].Value = fIva;
			
			if (MostrarIce)
			{
				stExec = string.Format("Exec ArticuloIce {0}, {1}", IdTipoFactura, stIdArt);
				double fIce = Funcion.dEscalarSQL(cdsSeteoF, stExec);
				e.Cell.Row.Cells["IceL"].Value = fIce;
			}
			#endregion Lee Iva

			#region Verifica si aplica revision
			double dCant = (double) e.Cell.Row.Cells["Cantidad"].Value;
			string stVerif = "Exec DetCompraVerifNegValida " + stIdArt + ", " + dCant.ToString("0.000000", us) + ", " + IdTipoFactura;
			if(!Funcion.bEjecutaSQL(cdsSeteoF, stVerif)) return true;
			#endregion verifica si aplica revision

			#region Verifica Negativos o alertas
			string stBodega = drCompra["Bodega"].ToString();
			if ((bool) drSeteoF["VerBodega"])
				stBodega = e.Cell.Row.Cells["Bodega"].Value.ToString();
			//bool FactManual = (bool)drCompra["Manual"];
			string stidDetC = e.Cell.Row.Cells["idDetCompra"].Value.ToString();
			stVerif = "Exec DetCompraVerifNegativo " + stIdArt + ", " + dCant.ToString("0.000000", us) + ", " + stBodega + ", " + stidDetC + ", " + IdTipoFactura + ", " + e.Cell.Row.Cells["Manual"].Value.ToString();
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

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.barraDato1.bEditar)
			{
				this.ultraGrid1.Enabled = true;
				ModificaGrilla(true);
				this.txtNotas.Enabled = true;			
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Bodega"].Value = 1;
			e.Row.Cells["Signo"].Value = IdSigno;
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!this.barraDato1.bEditar)return;

			if (!bEliminaAlValidar)
			{
				if (!this.barraDato1.bNuevo)
				{	
					for (int i = 0; i < e.Rows.Length; i++)
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
								MessageBox.Show("Este Articulo ya tiene las Seriales Registradas no puede Eliminarlo \n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
							}
							else
							{
								if (DialogResult.No ==	MessageBox.Show(string.Format("¿Esta Seguro de Borrar el Artículo '{0}', con {1} Seriales, \n\n Este proceso no puede ser revertido?", e.Rows[i].Cells["Codigo"].Value.ToString(), iConteo), "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
								{
									e.Cancel = true;
									return;
								}
								else
								{
									e.DisplayPromptMsg = false;

									string sSQLElimina = string.Format("Exec EliminaArticuloODCSeriales {0}, {1}, 11", (int)e.Rows[i].Cells["idDetCompra"].Value, iConteo);
									Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);
								}
							}
						}
						else
						{
							e.DisplayPromptMsg = false;

							string sSQLElimina = string.Format("Exec EliminaArticuloODCSeriales {0}, {1}, 11", (int)e.Rows[i].Cells["idDetCompra"].Value, iConteo);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);
						}
					}
				}
				else e.DisplayPromptMsg = false;
			}
			else e.DisplayPromptMsg = false;
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (drCompra == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			this.barraDato1.IdRegistro = (int)drCompra["idCompra"];
			this.barraDato1.ProximoId(5);

			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((int)drCompra["Estado"] == 9)
			{
				MessageBox.Show("No puede Anular una Reserva FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)drCompra["Estado"] == 5 && !miAcceso.BAnularAprobado)
			{
				MessageBox.Show("No puede Anular una Reserva APROBADA \n\n Consulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)drCompra["Estado"] == 3 && !miAcceso.BAnularProcesado)
			{
				MessageBox.Show("No puede Anular una Reserva PROCESADA \n\n Consulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sSQLValSeriales = string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, 11", (int)drCompra["idCompra"]);
			if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValSeriales, true) > 0 && !miAcceso.BAnularConSeriales)
			{
				MessageBox.Show("No puede Anular Existen Seriales Ingresadas \n\n Consulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		

			if (DialogResult.No == MessageBox.Show(string.Format("Esta Seguro de Anular Esta RESERVA '{0}'", this.lblEstado.Text.ToString()), "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;

			using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
			{
				if (DialogResult.OK == miNota.ShowDialog())
				{
					try
					{
						string sSQLAnul = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", (int)drCompra["idCompra"], IdTipoFactura, miNota.txtNotas.Text.ToString());
						Funcion.sEscalarSQL(cdsSeteoF, sSQLAnul);

						barraDato1.ProximoId(5);
						barraDato1.ToolTipBorrar(Funcion.Tiempo("Anular"));
						this.lblEstado.Text = "";
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}		
				}
			}
		}

		private void ReservaCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Deshacer(sender, e);
				}
				else
					Close();
			}
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.C)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.I)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Impresora(sender, e);
				}
			}
			if (e.KeyCode == Keys.Delete)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Borrar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drCompra["idCompra"];
					barraDato1.ProximoId(5);					
				}
			}
		}

		public void BuscarCompra(int id)
		{
			barraDato1.IdRegistro = id;
			barraDato1.ProximoId(5);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Buscar " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			using (frmBuscaTransacciones miBusqueda = new frmBuscaTransacciones(IdTipoFactura, 0))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{										
						barraDato1.IdRegistro = (int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
						barraDato1.ProximoId(5);
					}
					catch{}
				}
			}
		}

		#region Impresión
		private void ImprimeArchivo()
		{
			string stArchivo = Funcion.sEscalarSQL(cdsCompra, "Exec ImprimeFactura " + txtIdCompra.Value);
			StreamWriter sw = new StreamWriter("Imprimir.txt");
			sw.WriteLine(stArchivo);
			sw.Close();

			string stPuerto = "lpt";
			string stSelect = "Select Cast(IsNull(Valor, 0) as int) From SeteoG Where Codigo = 'FPVTISP'";
			int iPuerto = Funcion.iEscalarSQL(cdsCompra, stSelect);
			if (iPuerto < 10) 
			{
				if (iPuerto < 1) iPuerto = 1;
				stPuerto += iPuerto;
			}
			else
			{
				iPuerto-=10;
				stPuerto = "COM" + iPuerto;
			}
			Process p = new Process();
			ProcessStartInfo psi = new ProcessStartInfo("cmd.exe");
			psi.Arguments = "/c " +  @"type imprimir.txt > " + stPuerto;
			p.StartInfo = psi;
			p.Start(); 
		}

		private void Imprimir(int iTipo)
		{
			Funcion.TiempoInicia();
			#region Seguridad para imprimir y reimprimir y control de vacios
			DateTime dtIni = DateTime.Now;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.ReImprimir && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			// Cuando no hay registros
			if (Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From Compra Where IdTipoFactura = 11") < 2)// cmbTipoCompra.Value == System.DBNull.Value)
			{
				MessageBox.Show("No hay informacion para imprimir", "Información");
				return;
			}
			#endregion Seguridad para imprimir y reimprimir

			

			#region Autoimpresores Impresion
			if (bAutoImpresor && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede ReImprimir en AutoImpresores ", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Autoimpresores
			
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + drCompra["idCompra"].ToString(); // txtIdCompra.Value;
			if (bAutoImpresor) stReporte += ", 1";
			else stReporte += ", 0";
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			string stFiltro = "{Compra.idCompra} = " + drCompra["idCompra"].ToString();
			#endregion Nombre de Reporte

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				#region Reporte con procedimiento almacenada
				string stProc = "Exec SerieExisteTipo " + drCompra["idCompra"].ToString() + ", 'FPVTISP', " + 1;
				bool bProc = Funcion.bEjecutaSQL(cdsSeteoF, stProc);
				if (bProc == true) 
				{
					stFiltro = "";
					stReporte = stReporte.Substring(0, stReporte.Length - 4) + "SP.rpt";

					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@idCompra";
					discreteVal.Value = drCompra["idCompra"].ToString();
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
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + drCompra["idCompra"].ToString();
			Funcion.EjecutaSQL(cdsCompra, stGrabaImprime);
			drCompra["impreso"] = true;
			barraDato1.ToolTipImprimir(Funcion.Tiempo("Imprimir"));
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtIni);
			//lblTiempo.Text = ts.Seconds.ToString() + ":" + ts.Milliseconds.ToString();
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}
		#endregion Impresion

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			#region Numero de Copias de CompraNumero
			string stCopia = "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = " + IdTipoFactura.ToString();
			int iCopias = Funcion.iEscalarSQL(cdsCompraTabla, stCopia);
			if (iCopias < 1) iCopias = 1;
			#endregion Numero de Copias de CompraNumero
			#region Muestra pantalla con totales en grande
			if (IdTipoFactura == 1 && drSeteoF["TiempoVenta"] != System.DBNull.Value
				&& (int) drSeteoF["TiempoVenta"] > 0)
			{
				double dTotal = (double) txtTotal.Value;
				//double dRecibido = (double) spnVuelto.Value;
				CompraGracias miGracias = new CompraGracias((int) drSeteoF["TiempoVenta"], dTotal, 0);
				miGracias.MdiParent = this.MdiParent;
				miGracias.Show();
			}
			#endregion Muestra pantalla con totales
			for (int j=0;j<iCopias;j++)
			{
				if (IdTipoFactura == 1 && Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPVTISP'"))
					ImprimeArchivo();
				else
					Imprimir(iCopias);
			}
			#region da problemas porque se queda siempre activado la impresion no se porque pero asi se arregla
			mouse_event(0x0001, 0, 30, 0, 0);
			mouse_event(0x0002, 0, 0, 0, 0);
			mouse_event(0x0004, 0, 0, 0, 0);
			#endregion da problemas porque se queda siempre activado no se porque
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

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if ((bool) drCompra["Borrar"] == false)
			{
				MessageBox.Show("Solo puede Restaurar registros Anulados", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Recuperar)
			{
				MessageBox.Show("No tiene Acceso a Restaurar " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea Restaurar el Registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Information)) return;

			Funcion.TiempoInicia();
			string stExec = "Exec CompraRestaurar " + drCompra["idCompra"].ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stExec);
			barraDato1.ProximoId(5);
			barraDato1.ToolTipBorrar(Funcion.Tiempo("Restaurar"));
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoRuc_Leave(object sender, System.EventArgs e)
		{
			drCompra["idTipoRuc"] = (int) this.cmbTipoRuc.Value;

			if ((int)drCompra["idTipoRuc"] == 4) this.txtRuc.MaxLength = 13;
			else if ((int)drCompra["idTipoRuc"] == 5) this.txtRuc.MaxLength = 10;
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (this.barraDato1.bEditar)
			{				
				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, 1, {2}", this.txtRuc.Text.ToString(), IdTipoFactura, false);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				drC.Read();
				if (drC.HasRows)
				{
					drCompra["idCliente"] = (int) drC.GetValue(0);
					drCompra["RUC"] = (string) drC.GetValue(1);
					drCompra["Nombre"] = (string) drC.GetValue(2);
					MaximoCredito = (double) drC.GetValue(4);

					DateTime dtVence = (DateTime)drCompra["Fecha"];
					if (!miAcceso.CambiarFecha && miAcceso.BCambiarVencimiento)
					{
						drCompra["FechaVencimiento"] = dtVence.AddDays(8);
						this.dtVence.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
					}
					
					this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", (int)drCompra["idCliente"]));
				}
				drC.Close();
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Valida identificación
			string sSQL = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text, (int)drCompra["idTipoRuc"]);
			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				
			if (sMensaje.StartsWith("Error"))
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtRuc.Focus();
				return;
			}
			#endregion Valida identificación

			#region Verifica si el Cliente Existe
			sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, 1, {2}", this.txtRuc.Text.ToString(), IdTipoFactura, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drC.Read();
			if (this.barraDato1.bEditar && !drC.HasRows)
			{
				drC.Close();	
				MessageBox.Show("Cliente no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtRuc.Focus();
				e.Cancel = true;
				return;
			}				
			drC.Close();	
			#endregion Verifica si el Cliente Existe

			#region Valida si es Cliente de crédito
			if ((int)drCompra["idFormaPago"] == 6)
			{							
				int ClienteCredito = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select IsNull(ContadoCred, 0) From Cliente Where RUC = '{0}' And Proveedor = 0", drCompra["RUC"].ToString()));
				if (ClienteCredito == 1)
				{
					MessageBox.Show("El cliente no tiene Autorización para crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			}		
			#endregion Valida si es Cliente de crédito
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (miAcceso.BCambiarVencimiento) this.dtVence.Focus();
				else this.cmbVendedor.Focus();
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
			else if (this.barraDato1.bEditar && Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Reservas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			DateTime dtVence = (DateTime)this.dtFecha.Value;
			this.dtFechaEntrega.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			this.dtVence.CalendarInfo.MinDate = (DateTime)this.dtFecha.Value;
			drCompra["FechaVencimiento"] = dtVence.AddDays(8);
		}

		private void dtVence_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtVence.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtVence.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtVence_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraGrid1.Focus();
		}

		private void cmbVendedor_Validated(object sender, System.EventArgs e)
		{
			drCompra["idVendedor"] = (int)this.cmbVendedor.Value;	
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && ((bool) drSeteoF["ExigeVendedor"] == true && this.cmbVendedor.Text.ToString().Length == 0))
			{
				MessageBox.Show("Seleccione un vendedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbVendedor.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtFechaEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtpHoraEntrega.Focus();
		}

		private void dtFechaEntrega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.chkEntregado.Checked && this.dtFechaEntrega.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtFechaEntrega.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtpHoraEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				//if (this.dtpHoraEntrega.Value.ToString().Length > 0) drCompra["HoraEntrega"] = this.dtpHoraEntrega.Value.ToString();
				this.cmbDirecciones.Focus();
			}
		}

		private void dtpHoraEntrega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//			if (this.barraDato1.bEditar && this.chkEntregado.Checked && this.dtpHoraEntrega.Value == System.DBNull.Value)
//			{
//				MessageBox.Show("Ingrese la hora de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.dtpHoraEntrega.Focus();
//				e.Cancel = true;
//				return;
//			}
		}

		private void cmbDirecciones_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.chkEntregado.Checked && this.cmbDirecciones.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la dirección de entrega.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbDirecciones.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void cmbDirecciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.barraDato1.Focus();
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.chkEntregado.Focus();
		}

		private void dtpHoraEntrega_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkEntregado_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkEntregado_Leave(object sender, System.EventArgs e)
		{
			drCompra["Entregado"] = this.chkEntregado.Checked; 
		}

		private void cmbDirecciones_Validated(object sender, System.EventArgs e)
		{

		}

		private void dtFechaEntrega_Validated(object sender, System.EventArgs e)
		{
		
		}

		private void dtpHoraEntrega_Validated(object sender, System.EventArgs e)
		{

		}

		private void cmbDirecciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFechaEntrega_Leave(object sender, System.EventArgs e)
		{
			if (this.dtFechaEntrega.Value != System.DBNull.Value) drCompra["Entregado"] = 1; 
			drCompra["FechaEntrega"] = (DateTime)this.dtFechaEntrega.Value;
		}

		private void dtpHoraEntrega_Leave(object sender, System.EventArgs e)
		{
			if (this.dtpHoraEntrega.Value != System.DBNull.Value)
			{
				drCompra["Entregado"] = 1;
				//drCompra["HoraEntrega"] = (DateTime)this.dtpHoraEntrega.Value;
			}
		}

		private void cmbDirecciones_Leave(object sender, System.EventArgs e)
		{
			if (this.cmbDirecciones.Value != System.DBNull.Value)
			{
				drCompra["Entregado"] = 1;
				drCompra["idDireccion"] = (int)this.cmbDirecciones.Value;
			}
		}

		private void cmbVendedor_Enter(object sender, System.EventArgs e)
		{
			this.cmbVendedor.ToggleDropdown();
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (!this.barraDato1.bEditar)return;

			if (this.barraDato1.bNuevo)return;

			if (e.Cell.Column.ToString() == "Codigo" || e.Cell.Column.ToString() == "Articulo")
			{
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
								
				string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
				int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iConteo > 0)
				{
					MessageBox.Show("Este Articulo ya tiene las Seriales Registradas no puede Modificarlo \n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}				
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}

