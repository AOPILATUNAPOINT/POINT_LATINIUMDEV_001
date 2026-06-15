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
	/// Summary description for frmNotaDeDebitoVenta.
	/// </summary>
	public class frmNotaDeDebitoVenta : DevExpress.XtraEditors.XtraForm
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataRow drCompra;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		int IdTipoFactura = 27;
		bool bAsientoAutomatico = false;
		private Acceso miAcceso;
		bool bRastreo = false;
		private int IdSigno = 2;
		private bool bActualiza = true;
		private bool bIncluyeIva = false;
		bool bEliminaAlValidar = false;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label lblAnulado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFiscal;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.Misc.UltraButton btCreaAsiento;

		[DllImport("user32.dll")]
		static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNotaDeDebitoVenta()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetCompra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPeso");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaTelco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Elaboracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residuo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescGlobal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Razonable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CliFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoCDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PresupuestoC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PresupuestoCDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPasar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IceL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("costoBinaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("costoB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bordado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estampado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedioNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_IdArticulo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sugerido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescNC");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Manual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantAereo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantMaritimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Liquidado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoBarras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("unidadBotrosa", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetRubro", 2);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.btCreaAsiento = new Infragistics.Win.Misc.UltraButton();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.label54 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkFiscal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			this.SuspendLayout();
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
			this.barraDato1.Location = new System.Drawing.Point(8, 496);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 161;
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
			// btCreaAsiento
			// 
			this.btCreaAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.Image = 9;
			this.btCreaAsiento.Appearance = appearance1;
			this.btCreaAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCreaAsiento.Location = new System.Drawing.Point(312, 496);
			this.btCreaAsiento.Name = "btCreaAsiento";
			this.btCreaAsiento.Size = new System.Drawing.Size(24, 25);
			this.btCreaAsiento.TabIndex = 192;
			this.btCreaAsiento.Text = "A";
			this.btCreaAsiento.Click += new System.EventHandler(this.btCreaAsiento_Click);
			this.btCreaAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btCreaAsiento_MouseDown);
			// 
			// txtSerieFact
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance2;
			this.txtSerieFact.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SerieFactura"));
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Location = new System.Drawing.Point(136, 8);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(64, 22);
			this.txtSerieFact.TabIndex = 194;
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(232, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 193;
			this.label4.Text = "Número";
			// 
			// txtAutFactura
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance3;
			this.txtAutFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.AutFactura"));
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(88, 40);
			this.txtAutFactura.MaxLength = 10;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(320, 22);
			this.txtAutFactura.TabIndex = 199;
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(8, 40);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(65, 16);
			this.label31.TabIndex = 198;
			this.label31.Text = "Autorización";
			// 
			// txtNumero
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(296, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 197;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(640, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 195;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bodega"));
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(704, 9);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(128, 21);
			this.cmbBodega.TabIndex = 196;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// cmbTipoRuc
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance5;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idTipoRuc"));
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 233;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 112);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(136, 21);
			this.cmbTipoRuc.TabIndex = 206;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 112);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(86, 16);
			this.lblIdentificacion.TabIndex = 205;
			this.lblIdentificacion.Text = "N. Identificación";
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(392, 112);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(38, 16);
			this.lblCliente.TabIndex = 204;
			this.lblCliente.Text = "Cliente";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(392, 152);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 203;
			this.label6.Text = "Fecha";
			// 
			// txtRuc
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance6;
			this.txtRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Ruc"));
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(240, 112);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(144, 22);
			this.txtRuc.TabIndex = 201;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// dtFecha
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance7;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(472, 144);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 202;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			// 
			// txtNombre
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance8;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(472, 112);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(360, 22);
			this.txtNombre.TabIndex = 200;
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
			ultraGridColumn6.Header.VisiblePosition = 27;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Artículo";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 194;
			ultraGridColumn9.Header.Caption = "Unidad";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn9.Width = 53;
			ultraGridColumn10.Header.VisiblePosition = 33;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 30;
			ultraGridColumn11.Hidden = true;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn12.NullText = "0";
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 70;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 80;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Header.Caption = "Iva";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn15.Header.VisiblePosition = 64;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 60;
			ultraGridColumn16.Header.Caption = "Ref. Codigo";
			ultraGridColumn16.Header.VisiblePosition = 22;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 39;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.Header.Caption = "Serial";
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 150;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance13;
			ultraGridColumn18.Header.Caption = "Descuento";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 14;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance14;
			ultraGridColumn19.Header.Caption = "Peso";
			ultraGridColumn19.Header.VisiblePosition = 21;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 51;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance15;
			ultraGridColumn20.Header.VisiblePosition = 29;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn20.PromptChar = ' ';
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance16;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.MaskInput = "nnnnnnnnn";
			ultraGridColumn21.Width = 14;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance17;
			ultraGridColumn22.Header.Caption = "% Desc.";
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 65;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance18;
			ultraGridColumn23.Header.VisiblePosition = 26;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn23.PromptChar = ' ';
			ultraGridColumn24.Header.VisiblePosition = 28;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 39;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 31;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 32;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.Caption = "Centro Costo";
			ultraGridColumn30.Header.VisiblePosition = 34;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn31.Header.Caption = "Proyecto";
			ultraGridColumn31.Header.VisiblePosition = 19;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance19;
			ultraGridColumn32.Header.VisiblePosition = 7;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn32.NullText = "0";
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.UseEditorMaskSettings = true;
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 200;
			ultraGridColumn34.Header.VisiblePosition = 35;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance20;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 16;
			ultraGridColumn35.PromptChar = ' ';
			ultraGridColumn35.Width = 100;
			ultraGridColumn36.Header.Caption = "Importacion";
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.PromptChar = ' ';
			ultraGridColumn38.Header.VisiblePosition = 6;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 340;
			ultraGridColumn39.Header.Caption = "CentroCosto";
			ultraGridColumn39.Header.VisiblePosition = 15;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.Caption = "Proyecto";
			ultraGridColumn40.Header.VisiblePosition = 18;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 24;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 40;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.PromptChar = ' ';
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 44;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.PromptChar = ' ';
			ultraGridColumn48.Header.VisiblePosition = 45;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.PromptChar = ' ';
			ultraGridColumn49.Header.VisiblePosition = 46;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 47;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 48;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 49;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.Caption = "Vendedor";
			ultraGridColumn53.Header.VisiblePosition = 51;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn54.Header.Caption = "Obraid";
			ultraGridColumn54.Header.VisiblePosition = 50;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn55.Header.Caption = "Capituloid";
			ultraGridColumn55.Header.VisiblePosition = 52;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn56.Header.Caption = "Rubroid";
			ultraGridColumn56.Header.VisiblePosition = 54;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn57.Header.Caption = "Obra";
			ultraGridColumn57.Header.VisiblePosition = 53;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.MaskInput = "";
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn58.Header.VisiblePosition = 55;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.MaskInput = "";
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn59.Header.Caption = "Rubro";
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.MaskInput = "";
			ultraGridColumn60.Header.VisiblePosition = 62;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 57;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 60;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn63.Header.VisiblePosition = 56;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.MaskInput = "";
			ultraGridColumn64.Header.VisiblePosition = 61;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 23;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.PromptChar = ' ';
			ultraGridColumn65.Width = 43;
			ultraGridColumn66.Header.VisiblePosition = 63;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 65;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 66;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 70;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 67;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 47;
			ultraGridColumn72.Header.VisiblePosition = 69;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 50;
			ultraGridColumn73.Header.VisiblePosition = 71;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 48;
			ultraGridColumn74.Header.VisiblePosition = 72;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 44;
			ultraGridColumn75.Header.VisiblePosition = 73;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 68;
			ultraGridColumn76.Header.VisiblePosition = 74;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 50;
			ultraGridColumn77.Header.VisiblePosition = 75;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 60;
			ultraGridColumn78.Header.VisiblePosition = 81;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 82;
			ultraGridColumn79.Header.VisiblePosition = 88;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 82;
			ultraGridColumn80.Header.VisiblePosition = 76;
			ultraGridColumn80.Hidden = true;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance21;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 77;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn82.Header.VisiblePosition = 78;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 79;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 80;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 82;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 83;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 84;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 86;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 89;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 87;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.AutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			ultraGridColumn93.Header.Caption = "Código";
			ultraGridColumn93.Header.VisiblePosition = 5;
			ultraGridColumn93.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn93.Width = 214;
			ultraGridColumn94.Header.Caption = "Unidades";
			ultraGridColumn94.Header.VisiblePosition = 10;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn95.Header.VisiblePosition = 59;
			ultraGridColumn95.Hidden = true;
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
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 10F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance24;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 192);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(824, 136);
			this.ultraGrid1.TabIndex = 207;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idAsiento"));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(432, 496);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(48, 22);
			this.txtIdAsiento.TabIndex = 226;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Location = new System.Drawing.Point(384, 496);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(40, 22);
			this.txtIdCompra.TabIndex = 225;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(600, 456);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 224;
			this.label3.Text = "Total";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(600, 376);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(48, 16);
			this.lblSubTotal.TabIndex = 216;
			this.lblSubTotal.Text = "SubTotal";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(680, 432);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 215;
			this.lblIva.Text = "IVA";
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(600, 408);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(57, 16);
			this.lblDescuento.TabIndex = 214;
			this.lblDescuento.Text = "Descuento";
			// 
			// label54
			// 
			this.label54.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label54.AutoSize = true;
			this.label54.Location = new System.Drawing.Point(712, 352);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(32, 16);
			this.label54.TabIndex = 212;
			this.label54.Text = "IVA 0";
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label55.AutoSize = true;
			this.label55.Location = new System.Drawing.Point(808, 352);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 16);
			this.label55.TabIndex = 213;
			this.label55.Text = "IVA";
			this.label55.Visible = false;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance28;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Total"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(664, 448);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.ReadOnly = true;
			this.txtTotal.Size = new System.Drawing.Size(168, 34);
			this.txtTotal.TabIndex = 223;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance29;
			this.txtIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Iva"));
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(752, 424);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.ReadOnly = true;
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 222;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance30;
			this.txtDescIva0Total.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento0"));
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(664, 400);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.ReadOnly = true;
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 221;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance31;
			this.txtDescIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento"));
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(752, 400);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.ReadOnly = true;
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 220;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance32;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalIva"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(752, 376);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.ReadOnly = true;
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 219;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance33;
			this.txtIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalExcento"));
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(664, 376);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.ReadOnly = true;
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 218;
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(648, 152);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(112, 25);
			this.lblAnulado.TabIndex = 217;
			this.lblAnulado.Text = "ANULADO";
			// 
			// txtComprobante
			// 
			this.txtComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance34;
			this.txtComprobante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprobante"));
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(88, 352);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(128, 22);
			this.txtComprobante.TabIndex = 211;
			this.txtComprobante.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComprobante_KeyPress);
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(8, 352);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 210;
			this.lblComprobante.Text = "Comprobante";
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 384);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 209;
			this.label2.Text = "Notas";
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance35;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(88, 384);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(448, 96);
			this.txtNotas.TabIndex = 208;
			// 
			// chkFiscal
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFiscal.Appearance = appearance36;
			this.chkFiscal.BackColor = System.Drawing.Color.Transparent;
			this.chkFiscal.CausesValidation = false;
			this.chkFiscal.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Fiscal"));
			this.chkFiscal.Location = new System.Drawing.Point(8, 8);
			this.chkFiscal.Name = "chkFiscal";
			this.chkFiscal.Size = new System.Drawing.Size(48, 22);
			this.chkFiscal.TabIndex = 228;
			this.chkFiscal.Text = "Fiscal";
			this.chkFiscal.Validated += new System.EventHandler(this.chkFiscal_Validated);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 144);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 229;
			this.label1.Text = "Mótivo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMotivos
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance37;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idMotivo"));
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Enabled = false;
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(104, 144);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(248, 21);
			this.cmbMotivos.TabIndex = 230;
			this.cmbMotivos.ValueMember = "idMotivo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(88, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(29, 16);
			this.label5.TabIndex = 231;
			this.label5.Text = "Serie";
			// 
			// frmNotaDeDebitoVenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(840, 526);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbMotivos);
			this.Controls.Add(this.chkFiscal);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label54);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btCreaAsiento);
			this.Controls.Add(this.barraDato1);
			this.KeyPreview = true;
			this.Name = "frmNotaDeDebitoVenta";
			this.Text = "Nota de Débito Venta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNotaDeDebitoVenta_KeyDown);
			this.Load += new System.EventHandler(this.frmNotaDeDebitoVenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
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
				lblAnulado.Visible = (bool) drCompra["Borrar"];
				if ((int) e.Row["idTipoFactura"] < 1)	e.Row["idTipoFactura"] = 27;
			}
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

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void Seguridades()
		{
			#region Codigo de Seguridad
			string stTipo = "0418";
			string stPermiso = "60";
								
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			#endregion Tipo de Factura
			if (!Funcion.Permiso(stPermiso, cdsSeteo))
			{
				MessageBox.Show("No puede ingresar a Nota de Debito Compras", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			#region Cambiar Numeracion
			//			if (!miAcceso.CambiarNumeracion) 
			//			{
			//				txtNumero.ReadOnly = true;
			//			}
			#endregion Cambiar Numeracion
		} 
		private void Configuracion()
		{
			#region Asigna Signo
			string stSelect = "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 27";
			IdSigno = Funcion.iEscalarSQL(cdsCompra, stSelect);
			#endregion Asigna Signo
			barraDato1.stFiltroExt = " And idTipoFactura = 27";
			#region Lectura de Archivo - Propiedades - Configuracion
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#endregion Lectura de Archivo - Propiedades - Configuracion
		
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

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void frmNotaDeDebitoVenta_Load(object sender, System.EventArgs e)
		{
			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");
			Cursor = Cursors.Default;

			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;

			#region Ordenacion de grilla
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FVOG'")) ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			#endregion Ordenacion de grilla

			Seguridades();
			Configuracion();

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BodegaUsuario 27");
			this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 27");
			this.cmbMotivos.DisplayLayout.Bands[0].Columns["Motivo"].Width = 300;
			this.cmbMotivos.DisplayLayout.Bands[0].Columns["idMotivo"].Hidden = true;

			#region Auditoria			
			string stAudita = "Exec AuditaCrear 53, 6, 'NOTA DE DEBITO'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			#endregion Auditoria

			barraDato1.PosUltima();		

			#region Habilita controles
			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);
			
			#endregion habilita controles

			bAsientoAutomatico = Funcion.bEjecutaSQL(cdsCompra, "Exec AsientoAutomaticoVerifica " + IdTipoFactura);

			txtNotas.TabIndex = 1;
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (drCompra == null)
			{
				MessageBox.Show("No existe un registro para Anular", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
						
			if ((bool) lblAnulado.Visible == false)
			{				
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				string stExecAnula = "Exec CompraBorrarVerifica " + drCompra["idCompra"].ToString() + ", 0";
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExecAnula, true);
				if (stMensaje.StartsWith("No"))
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
				if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;
		
				lblAnulado.Visible = true;
				stExecAnula = "Exec CompraBorrar " + drCompra["idCompra"].ToString() + ", 0";
				stMensaje = Funcion.sEscalarSQL(cdsCompra, stExecAnula, true);
				
				if (stMensaje.Length == 0)
				{
					barraDato1.ProximoId(5);
					barraDato1.ToolTipBorrar(Funcion.Tiempo("Anular"));
				}
				else
					MessageBox.Show(stMensaje, "Error al Anular", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
			using (BuscaCompra miBusqueda = new BuscaCompra(IdTipoFactura))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{
						if (miBusqueda.chkOrden.Checked) barraDato1.DataOrden = miBusqueda.stOrden;
						else barraDato1.DataOrden = "";
						int id = int.Parse(miBusqueda.idBusca.Text);
						BuscarCompra(id);
					}
					catch{}
				}
			}
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

			barraDato1.CrearRegistro();
			
			ModificaGrilla(true);

			drCompra["idTipoFactura"] = 27;
			
			drCompra["idProyecto"] = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 27, 1");
			drCompra["Bodega"] = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 27, 3");
			drCompra["idComprobante"] = 5;
			drCompra["Fecha"] = DateTime.Today;
			drCompra["idTipoRuc"] = 5;
			drCompra["Fiscal"] = true;

			this.txtNombre.Enabled = false;			
			this.cmbBodega.Enabled = false;

			if (!miAcceso.CambiarFecha) this.dtFecha.Enabled = false;			

			this.txtSerieFact.Focus();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();

			this.btCreaAsiento.Focus();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			#region Verifica
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
			
			if (this.lblAnulado.Visible)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			

			if (drCompra["Bloquear"] != DBNull.Value && (bool) drCompra["Bloquear"] == true)
			{
				MessageBox.Show("No puede editar registro por estar bloqueado", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!miAcceso.EditaImpresa && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede Editar Facturas Impresas ", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (miAcceso.EditarSoloDia)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
				DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
				if (dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
				{
					MessageBox.Show("Solo puede editar Registros de fecha " + dtFechaHoy.ToString("dd/MMM/yyyy"), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			if (miAcceso.EditarDiaIngreso)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
				DateTime dtFechaIng = (DateTime) drCompra["FechaIngreso"];
				if (dtFechaHoy.ToShortDateString() != dtFechaIng.ToShortDateString())
				{	
					MessageBox.Show("Registro Ingresado el " +
						dtFechaIng.ToString("dd/MMM/yyyy") + 
						".\nSolo puede editar Asientos Ingresados el día de Hoy",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			#endregion Verifica

			barraDato1.EditarRegistro();

			ModificaGrilla(true);

			if (!(bool)drCompra["Fiscal"]) this.txtAutFactura.Enabled = false;
			else this.txtAutFactura.Enabled = true;

			barraDato1.ToolTipEditar(Funcion.Tiempo("Editar"));

			if (!miAcceso.CambiarFecha) this.dtFecha.Enabled = false;		
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			#region Verificaciones
			#region Cabecera
			#region Serie
			if (this.txtSerieFact.Text.ToString().Length == 0 && this.chkFiscal.Checked)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				return;
			}
			if (this.txtSerieFact.Text.ToString().Length < 6  && this.chkFiscal.Checked)
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
			
			string sSQL = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 28 And idCliente = 2 And SerieFactura = '{0}' And Numero = '{1}'", (int)drCompra["idCliente"], drCompra["SerieFactura"].ToString(), drCompra["Numero"].ToString());
			if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) > 0)
			{
				MessageBox.Show("El número ya existe con este número de serie y cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}			
			#endregion Numero
			#region Identificacion
			if (this.txtRuc.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de identificación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}

			string sSQLVal = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text, (int)drCompra["idTipoRuc"]);
			string SMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVal);
			if (SMensaje.StartsWith("Error"))
			{
				MessageBox.Show(SMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			
			string sSQLB = string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 0", this.txtRuc.Text.ToString());
			int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLB);
			if (iCount == 0)
			{
				MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			#endregion Identificacion
			#region Fecha
			if (this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#endregion Fecha
			#endregion Cabecera
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

					double dPrecio = (double) dr.Cells["Precio"].Value;
					
					if (dPrecio == 0 && dr.Cells["idArticulo"].Value != System.DBNull.Value) 
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
				MessageBox.Show("Ingrese el detalle de la nota de débito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}
			#endregion Verifica Ingreso en el grid
			#endregion Detalle
			#endregion Verificaciones

			cmdTotal();
			bool bNuevoReg = barraDato1.bNuevo;					

			#region Centro de costo
			string sSQL1 = string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)drCompra["Bodega"]);
			int IProyecto = Funcion.iEscalarSQL(cdsSeteoF, sSQL1);
			if (IProyecto > 0) drCompra["idProyecto"] = IProyecto;
			#endregion Centro de costo

			#region Grabar y mostrar
			barraDato1.Focus();
			if (!barraDato1.GrabaRegistro())
				if (bNuevoReg)
				{
					MessageBox.Show("No se puede grabar la factura sistema ocupado", "Información");
					return;
				}			
			ModificaGrilla(false);
			#endregion Grabar y mostrar			
						
			#region Actualiza Bodegas a las lineas
			if ((bool) drSeteoF["VerBodega"] == false)
			{
				string ActualizaBodega = "Exec DetCompraBodegaAct " + drCompra["idCompra"].ToString();
				Funcion.EjecutaSQL(cdsSeteoF, ActualizaBodega, true);
			}
			#endregion Actualiza Bodegas a las lineas
							
			barraDato1.ProximoId(5);
			
			ModificaGrilla(false);
			barraDato1.ToolTipNuevo(Funcion.Tiempo("Grabar"));
			
			bEliminaAlValidar = false;
			
			#region Asiento Automatico
			if (bAsientoAutomatico) AsientoGenerar();
			#endregion Asiento Automatico

			int iDFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDFPNotaDebito {0}", (int)this.txtIdCompra.Value));
			string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, 5, 1, {2}, 0, 0, '', '', '', '', '', 0, '', '', 0, 0, 0, 0", iDFP, (int)this.txtIdCompra.Value, (double)this.txtTotal.Value);							
			Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			#region Numero de Copias de CompraNumero
			string stCopia = "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = " + IdTipoFactura.ToString();
			int iCopias = Funcion.iEscalarSQL(cdsCompra, stCopia);
			if (iCopias < 1) iCopias = 1;
			#endregion Numero de Copias de CompraNumero
			
			for (int j=0;j<iCopias;j++)
			{
				Imprimir(iCopias);
			}
			#region da problemas porque se queda siempre activado la impresion no se porque pero asi se arregla
			mouse_event(0x0001, 0, 30, 0, 0);
			mouse_event(0x0002, 0, 0, 0, 0);
			mouse_event(0x0004, 0, 0, 0, 0);
			#endregion da problemas porque se queda siempre activado no se porque
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Imprimir(0);
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
			if (drCompra == null)
			{
				MessageBox.Show("No hay informacion para imprimir", "Información");
				return;
			}
			#endregion Seguridad para imprimir y reimprimir

			

			
			
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + drCompra["idCompra"].ToString();
			
			stReporte += ", 0";
						
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			string stFiltro = "{Compra.idCompra} = " + drCompra["idCompra"].ToString();
			#endregion Nombre de Reporte y Filtro

			#region Genera series en las lineas
			if (drSeteoF["SerieLinea"] != DBNull.Value && (bool) drSeteoF["SerieLinea"] == true)
			{
				string stExec = "Exec SerieLinea " + drCompra["idCompra"].ToString();
				Funcion.EjecutaSQL(cdsCompra, stExec, true);
			}
			#endregion Genera series en las lineas

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				#region Reporte con procedimiento almacenada
				string stProc = "Exec SerieExisteTipo " + drCompra["idCompra"].ToString() + ", 'FPVTISP', " + drCompra["idComprobante"].ToString();
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
					if (IdTipoFactura == 2)
					{
						string sSQLMail0 = string.Format("Select Count(*) From ClienteContacto Where EnvioMail = 1 And idCliente = {0}", (int)drCompra["idCliente"]);
						int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLMail0);
						if (iCount > 0)
						{
							miReporte.ExportaPDF(@"\OrdenCompra_" + drCompra["Numero"].ToString() + ".pdf");
							string Archivo = drCompra["Numero"].ToString() + ".pdf";//@"\\" + IP + @"\Latinium\OrdenCompra_ " + drCompra["Numero"].ToString() + ".pdf"; //"C:/Latinium/OrdenCompra_" + 
								
							string sSQLMail1 = string.Format(@"Exec MailOrdenCompra 'C:\Latinium\OrdenesDeCompra\OrdenCompra_{0}', {1}", Archivo, (int)drCompra["idCompra"]);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLMail1, true);
						}
					}
					else
					{
						miReporte.ImprimeRep(2);
					}
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + drCompra["idCompra"].ToString();
			Funcion.EjecutaSQL(cdsCompra, stGrabaImprime);
			drCompra["impreso"] = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCompra = cdsCompra.TableViews["Compra"].Rows[0];
				if ((bool)drCompra["Borrar"] == true) lblAnulado.Visible = true;
				else lblAnulado.Visible = false;
			}
			catch{};
			barraDato1.ToolTipBuscar(Funcion.Tiempo("Buscar"));
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
			string stExec = "Exec CompraRestaurar " + txtIdCompra.Value;
			Funcion.EjecutaSQL(cdsSeteoF, stExec);
			barraDato1.ProximoId(5);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Lista desplegable de Articulos
			if (drSeteoF["VerDesplegable"] != DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				e.Layout.Bands[0].Columns["Articulo"].Hidden = false;
				e.Layout.Bands[0].Columns["Codigo"].Hidden = false;
				e.Layout.Bands[0].Columns["CodArt"].Hidden = true;
				e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;
			}
			#endregion Lista desplegable de Articulo
			#region Unidades
			e.Layout.Bands[0].Columns["idUnidad"].Hidden = !(bool) drSeteoF["VerUnidad"];
			e.Layout.Bands[0].Columns["Unidades"].Hidden = !(bool) drSeteoF["VerUnidad"];
			if ((bool) drSeteoF["VerUnidad"]) e.Layout.Bands[0].Columns["Cantidad"].CellActivation = Activation.Disabled;
			#endregion Unidades
			if (!(bool)drSeteoF["CambiarIva"]) e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;
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
			#endregion Decimales para el total
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
			#endregion Decimales para la cantidad, peso y Pedido
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
			#endregion Decimales para el Impuesto
			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPorcDesc"]);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			e.Layout.Bands[0].Columns["DescNC"].Format = stFormato;
			
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, iDecPorD);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["DescNC"].MaskInput = stInput;			
			#endregion Decimales para porcentaje de descuento
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
						if (bRastreo) Rastreo("Cambio de CodArt Asigna idArticulo");
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
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, 1, 0", stCod, IdTipoFactura, (int) this.cmbBodega.Value);
				if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				#region Mas de una coincidencia o ninguna muestra la lista
				if (iCuenta != 1)
				{ 
					if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));
				
					#region Busqueda servicios
					if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && barraDato1.bEditar)
					{							
						using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(stCod, 27))
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								IdArt = miBusqueda.idArticulo;
								ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
							}
						}
					}										
					#endregion Busqueda servicios
				}				
					#endregion Mas de una coincidencia o ninguna muestra la lista
					#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
				else 
				{
					if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
					stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, (int) this.cmbBodega.Value, IdTipoFactura);
					IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
				}
				#endregion Encuentra un solo articulo

				#region Asignaciones
				if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
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
			
			#region Verificaciones de cambio de articulo
			if (VerificaCambio > 0)
			{				
				cmdTotal();
			}
			#endregion Verificaciones
			
			#endregion Fin de Verificacion de Cambio
			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				bActualiza = false;
								
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

				double dDesc = dCantidad * dPrecio * (double) e.Cell.Row.Cells["DescuentoPorc"].Value / 100.0;
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
				cmdTotal();
			}
			#endregion Precio

			bActualiza = true;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Bodega"].Value = (int) this.cmbBodega.Value;
			e.Row.Cells["Signo"].Value = IdSigno;
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			int iTotalLineas = ultraGrid1.Rows.Count;
			string stExec = string.Format("Exec CompraNumeroLineas {0}, {1}", IdTipoFactura, iTotalLineas);
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);

			if (stMensaje.Length > 0) 
			{
				e.Cancel = true;
				MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information); // "No puede crear más lineas");
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
				string stMensaje = "Error Grilla: Nota de débito";
				int iLong = stMensaje.Length;
				if (iLong > 20) iLong = 20;

				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);

				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol
					+ "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", 
					"Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["CodArt"].Value = e.Row.Cells["IdArticulo"].Value;
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
				string stExec = string.Format("Exec ArticuloUnidadCambio {0}, {1}", IdArticulo, IdUnidad);
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
				e.Cell.Row.Cells["Cantidad"].Value = (double) e.Cell.Row.Cells["Unidades"].Value * dCant;
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
						e.Cell.Row.Cells["Precio"].Value = (double) e.Cell.Row.Cells["Precio"].Value / dCantidad;
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
			double dIverde = 0;

			#region Calculo Forma Pago Mixto 
			double dSubTotal1 = 0;
			double dSubTotal2 = 0;
			double dSubTotal3 = 0;
			double dSubTotal4 = 0;
			#endregion Calculo Forma Pago Mixto 

			#region Tempocodeca
			double dEstampado = 0;
			double dBordado = 0;
			double dOtros = 0;
			#endregion Tempocodeca
			
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
				
				#region Tempocodeca
				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTCEBO'")) 				
				{
					if (dr.Cells["Estampado"].Value != DBNull.Value) dEstampado = (double) dr.Cells["Estampado"].Value;
					if (dr.Cells["Bordado"].Value != DBNull.Value)	dBordado = (double) dr.Cells["Bordado"].Value;
					if (dr.Cells["Otros"].Value != DBNull.Value)	dOtros = (double) dr.Cells["Otros"].Value;
					dSubTotal = dCantidad * (dPrecio + dEstampado + dBordado + dOtros) + .000001;
				}
				else	
				{
					#endregion Tempocodeca
					dSubTotal = dCantidad * dPrecio + .000000001;
				}

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
			dSubtIva0 = Math.Round(dSubtIva0, dRedondeo);
			dSubtIva = Math.Round(dSubtIva, dRedondeo);
		
			drCompra["SubTotalExcento"] = dSubtIva0;
			drCompra["SubTotalIva"] = dSubtIva;
			if ((bool) drSeteoF["IncluyeIva"])
				drCompra["SubTotalIva"] = dSubtIva;
			else
				drCompra["SubTotalIva"] = dSubtIva;

			
			drCompra["Iva"] = dImpuesto;
			
			#endregion Calculo de subtotales

			#region Calculo del Total
			double dTotal = 0;
			dTotal = dSubtIva0 + dSubtIva + dImpuesto - dDescuento; 
			dTotal = Math.Round(dTotal, dRedondeo);

			if(dTotal < 0)	
			{
				drCompra["Total"] = 0;
				MessageBox.Show("Error en el calculo del total: " + dTotal.ToString("#,##0.00"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				drCompra["Total"] = dTotal;
			}
			#endregion Calculo del Total 
		}

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
			return cadena1;
		}

		private void AsientoAbrir()
		{
			if (txtIdAsiento.Value == DBNull.Value || (int) txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe un asiento creado de esta transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int IdAsiento = (int) txtIdAsiento.Value;
			if (drCompra["idAsientoPer"] != DBNull.Value && (int) drCompra["idAsientoPer"] > 0)
			{
				MessageBox.Show("Origen es un Asiento de Comprobante", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				IdAsiento = (int) drCompra["idAsientoPer"];
			}
			
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void AsientoGenerar()
		{
			#region Verificaciones
			// Crea asiento de Factura
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (barraDato1.bEditar) 
			{
				MessageBox.Show("No puede crear un asiento mientras edita la factura");
				return;
			}
			if ((bool) drCompra["Borrar"]) 
			{
				MessageBox.Show("No puede crear un asiento de una transacción Anulada");
				return;
			}

			if (drSeteoC["FechaFin"] != DBNull.Value && drSeteoC["FechaIni"] != DBNull.Value)
			{
				DateTime dtFechaAsi = (DateTime)dtFecha.Value;
				DateTime dtFechaAFin = (DateTime)drSeteoC["FechaFin"];
				DateTime dtFechaAIni = (DateTime)drSeteoC["FechaIni"];
				if (dtFechaAsi > dtFechaAFin || dtFechaAsi < dtFechaAIni)
				{
					MessageBox.Show("Fecha debe estar entre el " 
						+ dtFechaAIni.ToString("dd/MMM/yyyy")+ " y el " 
						+ dtFechaAFin.ToString("dd/MMM/yyyy") + ".\n"
						+ "Fechas definidas en Archivo - Propiedades - Contabilidad - Pestaña General", "Información");
					return;
				}
			}
			#endregion Verificaciones
			#region Verifica si la transaccion genera Asiento
			string stProcedimiento = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraAsientoTipo " + drCompra["idCompra"].ToString(), true);

			if (stProcedimiento.Length == 0)
			{
				MessageBox.Show("Tipo de Transaccion no genera Asiento", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Verifica si la transaccion genera Asiento
			#region Confirmacion
			if (btCreaAsiento.Visible)
				if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			this.Cursor = Cursors.WaitCursor;
			#endregion Confirmacion
			#region verificacion de Existencia
			string stExisteAsiento = "Exec AsientoMaestroExiste "	+ drCompra["idCompra"].ToString();
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				Cursor = Cursors.Default;
				return;
			}
			stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 2", true);
			if (stMensaje.Length > 0)
			{
				if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
				{
					this.Cursor = Cursors.Default;
					return;
				}
				stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 3", true);
			}
			#endregion verificacion
			#region Crea Asiento
			string stGeneraCompra = "EXEC " + stProcedimiento + drCompra["idCompra"].ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);
			MessageBox.Show("Asiento Generado con éxito", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			#endregion Crea Asiento
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
						#region F3
					case (int) Keys.F3:
						#region Articulos
						string sArticulo = ""; 
						sArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						int IdArt = 0;
												
						if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO") && barraDato1.bEditar)
						{							
							using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(sArticulo, 27))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									IdArt = miBusqueda.idArticulo;
									ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
								}
							}
						}									
									
						#endregion Articulos
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
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "REFNUMERO")
			{
				if (!Char.IsSeparator(e.KeyChar)) e.Handled = false; 
				else e.Handled = true;
			}

			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRCxG'")) // cambia la comilla por guion				
			{		
				try
				{					
					if(e.KeyChar == 39 )//&& e.KeyChar = guion)
					{						
						e.Handled = true;
						SendKeys.Send("-");											
					}
				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.Message, "Error");
				}
			}
		}

		private void btCreaAsiento_Click(object sender, System.EventArgs e)
		{
			AsientoGenerar();
		}

		private void btCreaAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			AsientoAbrir();
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void chkFiscal_Validated(object sender, System.EventArgs e)
		{
			if (!(bool)this.chkFiscal.CheckedValue) this.txtAutFactura.Enabled = false;
			else this.txtAutFactura.Enabled = true;
		}

		private void txtSerieFact_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void txtSerieFact_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
      if (this.chkFiscal.Checked && this.barraDato1.bEditar && this.txtSerieFact.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;
				return;
			}
			if (this.chkFiscal.Checked && this.barraDato1.bEditar && this.txtSerieFact.Text.ToString().Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtAutFactura.Focus();
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			char Pad = '0';
			drCompra["Numero"] = this.txtNumero.Value.ToString().PadLeft(9, Pad);
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.txtNumero.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				e.Cancel = true;
				return;
			}
			if (this.barraDato1.bEditar)
			{
				string sSQL = string.Format("Select COUNT(*) From Compra Where idTipoFactura = 27 And idCliente = 2 And SerieFactura = '{0}' And Numero = '{1}'", (int)drCompra["idCliente"], drCompra["SerieFactura"].ToString(), drCompra["Numero"].ToString());
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) > 0)
				{
					MessageBox.Show("El número ya existe con este número de serie y cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void txtAutFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), IdTipoFactura, (int)drCompra["idFormaPago"], (int)drCompra["idTipoRuc"], 0, (bool)drCompra["Consignacion"]))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						drCompra["idCliente"] = miBusqueda.idCliente;
						drCompra["Nombre"] = miBusqueda.Nombre;
						drCompra["Ruc"] = miBusqueda.Ruc;
						drCompra["idTipoRuc"] = miBusqueda.idTipoRuc;
					}
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13)
			{
				if (this.dtFecha.Enabled) this.dtFecha.Focus();
				else this.ultraGrid1.Focus();
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			#region Busca Cliente - Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), IdTipoFactura, (int)this.cmbBodega.Value, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (drC.HasRows)
			{
				drCompra["idCliente"] = (int) drC.GetValue(0);
				drCompra["Ruc"] = (string) drC.GetValue(1);
				drCompra["Nombre"] = (string) drC.GetValue(2);
			}
			drC.Close();
			#endregion Busca Cliente - Proveedor
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.txtRuc.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de identificación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				e.Cancel = true;
				return;
			}

			if (this.barraDato1.bEditar)
			{
				string sSQL = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text, (int)drCompra["idTipoRuc"]);
				string SMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
				
				if (SMensaje.StartsWith("Error"))
				{
					MessageBox.Show(SMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			
				string sSQLB = string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 0", this.txtRuc.Text.ToString());
				int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQLB);
				if (iCount == 0)
				{
					MessageBox.Show("Cliente no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtComprobante_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void frmNotaDeDebitoVenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
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

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

