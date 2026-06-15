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

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Compra2.
	/// </summary>
	public class Compra : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvComprobante;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label lblAnulado;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataView cdvNumero;
		private C1.Data.C1DataSet cdsCompraTabla;
		public DevExpress.XtraEditors.SimpleButton btTotal;
		private Infragistics.Win.UltraWinGrid.UltraDropDown CmbDescCod;
		private Infragistics.Win.UltraWinGrid.UltraDropDown CmbCodDesc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDias;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtVence;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnBodega;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCompra;
		private Infragistics.Win.Misc.UltraButton btSaldo1;
		private System.Windows.Forms.ToolTip toolTip;
		private Infragistics.Win.Misc.UltraButton btTransforma;
		private Infragistics.Win.Misc.UltraButton btNotaEntrega;
		private Infragistics.Win.Misc.UltraButton btBuscaArticulo;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private Infragistics.Win.Misc.UltraButton btCreaAsiento;
		private Infragistics.Win.Misc.UltraButton btGuiaRemision;
		private Infragistics.Win.Misc.UltraButton btActualizaPrecio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBorrar;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgContadoCred;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescPorcIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescPorcIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescLinea0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescLinea;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFlete;
		private Infragistics.Win.Misc.UltraButton btRefresaI;
		private Infragistics.Win.Misc.UltraButton btRefrescaD;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRevisado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRevision;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteresPagar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEmpezarEn;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPagosCada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcInteres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnticipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.Misc.UltraButton btTotalCuotas;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIceTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCuenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCredTribut;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet OptBienServicio;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TabPageCOA;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDevolucion;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutImprenta;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPrioridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutPrincipal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutPresupuestaria;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConvenioInt;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRefrendo;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.Misc.UltraButton cmdVuelto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnVuelto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDepartamento;
		private System.Windows.Forms.Label label37;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtro;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecibe;
		private Infragistics.Win.Misc.UltraButton btFlete;
		private Infragistics.Win.Misc.UltraButton btNuevo;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.Misc.UltraButton btCambiarPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrBusca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btIngresar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCodBarras;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbCaducidad;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUnidad;
		private C1.Data.C1DataSet cdsUnidad;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbIce;
		private Infragistics.Win.Misc.UltraButton btComprobante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprob1;
		private Infragistics.Win.Misc.UltraButton btComprob1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBulto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPeso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private System.Windows.Forms.Label label43;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.Misc.UltraButton btImportacion;
		private Infragistics.Win.Misc.UltraButton btDescuento;
		private C1.Data.C1DataSet cdsTipoImportacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbImportacion;
		private System.Windows.Forms.Label label44;
		private Infragistics.Win.Misc.UltraButton btBarras;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCIF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tbpAutorizacion;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcTransaccion;
		private Infragistics.Win.Misc.UltraButton btCentro;
		private Infragistics.Win.Misc.UltraButton btSubProyecto;
		private C1.Data.C1DataView cdvEntrega;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtServicio;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label lblServicio;
		private System.Windows.Forms.Label lblPropina;
		private Infragistics.Win.Misc.UltraButton btComision;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.Misc.UltraButton btCorresponsal;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		private System.Windows.Forms.Label lblTiempo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodegaLinea;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPropina;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUnidadB;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecibeFabara;
		private C1.Data.C1DataSet cdsConfiteca;
		private C1.Data.C1DataView cdvConfiteca;
		private System.Windows.Forms.Label lblComisionExt;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		public DevExpress.XtraEditors.SimpleButton btnComisionExt;
		private Infragistics.Win.Misc.UltraButton btVendedor;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.Misc.UltraButton btFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumFrecCopiarCotiz;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumCopiarCotiz;
		public DevExpress.XtraEditors.SimpleButton btnCopiarCotiz;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedorGrilla;
		private DevExpress.XtraEditors.SimpleButton btDigitalizacion;
		private System.ComponentModel.IContainer components;

		[DllImport("user32.dll")]
		static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

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
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PorcIce", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSri");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCredTribut", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_ConfitecaComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorccomisionRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referido");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ImportacionTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Compra));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetCompra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "CmbDescCod");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad", -1, "cmbUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPeso");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbSubProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion", -1, "cmbImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaTelco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Elaboracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residuo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescGlobal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Razonable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CliFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumFabara");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor", -1, "cmbVendedorGrilla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt", 0, "CmbCodDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("unidadBotrosa", 1, "cmbUnidadB");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComprobante", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloUnidad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloUnidad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btVendedor = new Infragistics.Win.Misc.UltraButton();
			this.btSubProyecto = new Infragistics.Win.Misc.UltraButton();
			this.btCentro = new Infragistics.Win.Misc.UltraButton();
			this.btComprobante = new Infragistics.Win.Misc.UltraButton();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label15 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCambiarPrecio = new Infragistics.Win.Misc.UltraButton();
			this.spnVuelto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmdVuelto = new Infragistics.Win.Misc.UltraButton();
			this.label35 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtNumCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.rdgContadoCred = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.btFormaPago = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btCorresponsal = new Infragistics.Win.Misc.UltraButton();
			this.btComision = new Infragistics.Win.Misc.UltraButton();
			this.btDescuento = new Infragistics.Win.Misc.UltraButton();
			this.btFlete = new Infragistics.Win.Misc.UltraButton();
			this.txtDescPorcIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescPorcIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescLinea0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescLinea = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFlete = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btRefresaI = new Infragistics.Win.Misc.UltraButton();
			this.btRefrescaD = new Infragistics.Win.Misc.UltraButton();
			this.label17 = new System.Windows.Forms.Label();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbRecibeFabara = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btComprob1 = new Infragistics.Win.Misc.UltraButton();
			this.txtComprob1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btImportacion = new Infragistics.Win.Misc.UltraButton();
			this.label38 = new System.Windows.Forms.Label();
			this.cmbRecibe = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEntrega = new C1.Data.C1DataView();
			this.label37 = new System.Windows.Forms.Label();
			this.txtOtro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDepartamento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkRevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaRevision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label20 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtInteresPagar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEmpezarEn = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPagosCada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btTotalCuotas = new Infragistics.Win.Misc.UltraButton();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.TabPageCOA = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label40 = new System.Windows.Forms.Label();
			this.cmbIce = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.chkConvenioInt = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDevolucion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.OptBienServicio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtRefrendo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbCredTribut = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.tbpAutorizacion = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbCaducidad = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.chkAutPresupuestaria = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAutPrincipal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.spnPrioridad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.txtAutImprenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkCodBarras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btIngresar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugrBusca = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.spnNumFrecCopiarCotiz = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnNumCopiarCotiz = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCopiarCotiz = new DevExpress.XtraEditors.SimpleButton();
			this.btnComisionExt = new DevExpress.XtraEditors.SimpleButton();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvConfiteca = new C1.Data.C1DataView();
			this.cdsConfiteca = new C1.Data.C1DataSet();
			this.lblComisionExt = new System.Windows.Forms.Label();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.label45 = new System.Windows.Forms.Label();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCIF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label44 = new System.Windows.Forms.Label();
			this.cmbImportacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsTipoImportacion = new C1.Data.C1DataSet();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label43 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label42 = new System.Windows.Forms.Label();
			this.txtPeso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.label41 = new System.Windows.Forms.Label();
			this.txtBulto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cdvNumero = new C1.Data.C1DataView();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btTotal = new DevExpress.XtraEditors.SimpleButton();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.CmbDescCod = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.CmbCodDesc = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.spnDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtVence = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.spnBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbComprobante1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoCompra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btSaldo1 = new Infragistics.Win.Misc.UltraButton();
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			this.btBuscaArticulo = new Infragistics.Win.Misc.UltraButton();
			this.btTransforma = new Infragistics.Win.Misc.UltraButton();
			this.btNotaEntrega = new Infragistics.Win.Misc.UltraButton();
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btCreaAsiento = new Infragistics.Win.Misc.UltraButton();
			this.btGuiaRemision = new Infragistics.Win.Misc.UltraButton();
			this.btActualizaPrecio = new Infragistics.Win.Misc.UltraButton();
			this.btNuevo = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBorrar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.utcTransaccion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIceTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbSubProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.cmbUnidad = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsUnidad = new C1.Data.C1DataSet();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBarras = new Infragistics.Win.Misc.UltraButton();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.txtServicio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblServicio = new System.Windows.Forms.Label();
			this.lblPropina = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblTiempo = new System.Windows.Forms.Label();
			this.cmbBodegaLinea = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtPropina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbUnidadB = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cmbVendedorGrilla = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btDigitalizacion = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnVuelto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadoCred)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFlete)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibeFabara)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprob1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRevision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteresPagar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpezarEn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagosCada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			this.TabPageCOA.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbIce)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.OptBienServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefrendo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).BeginInit();
			this.tbpAutorizacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPrioridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrBusca)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.spnNumFrecCopiarCotiz)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumCopiarCotiz)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvConfiteca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfiteca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBulto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbDescCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbCodDesc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).BeginInit();
			this.utcTransaccion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIceTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPropina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadB)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedorGrilla)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btVendedor);
			this.ultraTabPageControl1.Controls.Add(this.btSubProyecto);
			this.ultraTabPageControl1.Controls.Add(this.btCentro);
			this.ultraTabPageControl1.Controls.Add(this.btComprobante);
			this.ultraTabPageControl1.Controls.Add(this.cmbVendedor);
			this.ultraTabPageControl1.Controls.Add(this.txtComprobante);
			this.ultraTabPageControl1.Controls.Add(this.cmbSubProyecto);
			this.ultraTabPageControl1.Controls.Add(this.cmbProyecto);
			this.ultraTabPageControl1.Controls.Add(this.label15);
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(486, 115);
			// 
			// btVendedor
			// 
			this.btVendedor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btVendedor.Location = new System.Drawing.Point(8, 64);
			this.btVendedor.Name = "btVendedor";
			this.btVendedor.Size = new System.Drawing.Size(64, 23);
			this.btVendedor.TabIndex = 94;
			this.btVendedor.Text = "Vendedor";
			this.btVendedor.Click += new System.EventHandler(this.btVendedor_Click);
			// 
			// btSubProyecto
			// 
			this.btSubProyecto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubProyecto.Location = new System.Drawing.Point(240, 88);
			this.btSubProyecto.Name = "btSubProyecto";
			this.btSubProyecto.Size = new System.Drawing.Size(72, 23);
			this.btSubProyecto.TabIndex = 93;
			this.btSubProyecto.Text = "Proyecto";
			this.btSubProyecto.Click += new System.EventHandler(this.btSubProyecto_Click);
			this.btSubProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSubProyecto_MouseUp);
			// 
			// btCentro
			// 
			this.btCentro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCentro.Location = new System.Drawing.Point(8, 88);
			this.btCentro.Name = "btCentro";
			this.btCentro.Size = new System.Drawing.Size(64, 23);
			this.btCentro.TabIndex = 92;
			this.btCentro.Text = "Centro C.";
			this.btCentro.Click += new System.EventHandler(this.btCentro_Click);
			this.btCentro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCentro_MouseUp);
			// 
			// btComprobante
			// 
			this.btComprobante.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btComprobante.Location = new System.Drawing.Point(448, 64);
			this.btComprobante.Name = "btComprobante";
			this.btComprobante.Size = new System.Drawing.Size(24, 23);
			this.btComprobante.TabIndex = 91;
			this.btComprobante.Text = "...";
			this.btComprobante.Click += new System.EventHandler(this.btComprobante_Click);
			this.btComprobante.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btComprobante_MouseUp);
			// 
			// cmbVendedor
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance1;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idVendedor"));
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(75, 64);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(149, 21);
			this.cmbVendedor.TabIndex = 90;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbVendedor_MouseUp);
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
			// txtComprobante
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance2;
			this.txtComprobante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprobante"));
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(320, 64);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(120, 21);
			this.txtComprobante.TabIndex = 89;
			// 
			// cmbSubProyecto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubProyecto.Appearance = appearance3;
			this.cmbSubProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idSubProyecto"));
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 337;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto.Location = new System.Drawing.Point(320, 88);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(152, 21);
			this.cmbSubProyecto.TabIndex = 88;
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbSubProyecto_MouseUp);
			// 
			// cmbProyecto
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance4;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idProyecto"));
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 408;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(75, 88);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(149, 21);
			this.cmbProyecto.TabIndex = 87;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Validating += new System.ComponentModel.CancelEventHandler(this.cmbProyecto_Validating);
			this.cmbProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbProyecto_MouseUp);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(240, 64);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(73, 16);
			this.label15.TabIndex = 85;
			this.label15.Text = "Comprobante";
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance5;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(8, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(466, 48);
			this.txtNotas.TabIndex = 82;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.txtComision);
			this.ultraTabPageControl2.Controls.Add(this.label1);
			this.ultraTabPageControl2.Controls.Add(this.cmbPrecio);
			this.ultraTabPageControl2.Controls.Add(this.btCambiarPrecio);
			this.ultraTabPageControl2.Controls.Add(this.spnVuelto);
			this.ultraTabPageControl2.Controls.Add(this.cmdVuelto);
			this.ultraTabPageControl2.Controls.Add(this.label35);
			this.ultraTabPageControl2.Controls.Add(this.label24);
			this.ultraTabPageControl2.Controls.Add(this.txtNumCuenta);
			this.ultraTabPageControl2.Controls.Add(this.label18);
			this.ultraTabPageControl2.Controls.Add(this.txtCheque);
			this.ultraTabPageControl2.Controls.Add(this.cmbFormaPago);
			this.ultraTabPageControl2.Controls.Add(this.rdgContadoCred);
			this.ultraTabPageControl2.Controls.Add(this.btPago);
			this.ultraTabPageControl2.Controls.Add(this.btFormaPago);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(486, 115);
			// 
			// txtComision
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComision.Appearance = appearance6;
			this.txtComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comision"));
			this.txtComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(312, 56);
			this.txtComision.MaskInput = "nnnnn.nn";
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(104, 21);
			this.txtComision.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtComision.TabIndex = 103;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(256, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(52, 16);
			this.label1.TabIndex = 102;
			this.label1.Text = "Comisión";
			// 
			// cmbPrecio
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio.Appearance = appearance7;
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio.DisplayLayout.AutoFitColumns = true;
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio.Location = new System.Drawing.Point(256, 80);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(128, 21);
			this.cmbPrecio.TabIndex = 100;
			this.toolTip.SetToolTip(this.cmbPrecio, "Escoja Tipo de Precio al que desea modificar");
			this.cmbPrecio.ValueMember = "idTipoPrecio";
			this.cmbPrecio.Visible = false;
			// 
			// btCambiarPrecio
			// 
			this.btCambiarPrecio.Location = new System.Drawing.Point(392, 80);
			this.btCambiarPrecio.Name = "btCambiarPrecio";
			this.btCambiarPrecio.Size = new System.Drawing.Size(24, 24);
			this.btCambiarPrecio.TabIndex = 99;
			this.btCambiarPrecio.Text = "$";
			this.btCambiarPrecio.Visible = false;
			this.btCambiarPrecio.Click += new System.EventHandler(this.btCambiarPrecio_Click);
			this.btCambiarPrecio.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCambiarPrecio_MouseDown);
			// 
			// spnVuelto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnVuelto.Appearance = appearance8;
			this.spnVuelto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Vuelto"));
			this.spnVuelto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnVuelto.FormatString = "#,##0.00";
			this.spnVuelto.Location = new System.Drawing.Point(80, 80);
			this.spnVuelto.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnVuelto.Name = "spnVuelto";
			this.spnVuelto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnVuelto.PromptChar = ' ';
			this.spnVuelto.Size = new System.Drawing.Size(88, 21);
			this.spnVuelto.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnVuelto.TabIndex = 98;
			// 
			// cmdVuelto
			// 
			this.cmdVuelto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmdVuelto.Location = new System.Drawing.Point(192, 80);
			this.cmdVuelto.Name = "cmdVuelto";
			this.cmdVuelto.Size = new System.Drawing.Size(48, 23);
			this.cmdVuelto.TabIndex = 57;
			this.cmdVuelto.Text = "Vuelto";
			this.cmdVuelto.Click += new System.EventHandler(this.cmdVuelto_Click);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(16, 80);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(52, 16);
			this.label35.TabIndex = 56;
			this.label35.Text = "Recibido:";
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(256, 32);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(44, 16);
			this.label24.TabIndex = 54;
			this.label24.Text = "Cuenta:";
			// 
			// txtNumCuenta
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuenta.Appearance = appearance9;
			this.txtNumCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCuenta"));
			this.txtNumCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuenta.Location = new System.Drawing.Point(312, 32);
			this.txtNumCuenta.Name = "txtNumCuenta";
			this.txtNumCuenta.Size = new System.Drawing.Size(104, 21);
			this.txtNumCuenta.TabIndex = 53;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(256, 8);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 16);
			this.label18.TabIndex = 52;
			this.label18.Text = "Cheque:";
			// 
			// txtCheque
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance10;
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(312, 8);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(104, 21);
			this.txtCheque.TabIndex = 51;
			// 
			// cmbFormaPago
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance11;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idFormaPago"));
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(16, 48);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(152, 21);
			this.cmbFormaPago.TabIndex = 50;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Validated += new System.EventHandler(this.cmbFormaPago_Validated);
			// 
			// rdgContadoCred
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgContadoCred.Appearance = appearance12;
			this.rdgContadoCred.BackColor = System.Drawing.Color.Transparent;
			this.rdgContadoCred.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ContadoCredito"));
			this.rdgContadoCred.ItemAppearance = appearance13;
			this.rdgContadoCred.ItemOrigin = new System.Drawing.Point(10, 3);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Contado";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Crédito";
			this.rdgContadoCred.Items.Add(valueListItem1);
			this.rdgContadoCred.Items.Add(valueListItem2);
			this.rdgContadoCred.Location = new System.Drawing.Point(16, 16);
			this.rdgContadoCred.Name = "rdgContadoCred";
			this.rdgContadoCred.Size = new System.Drawing.Size(152, 24);
			this.rdgContadoCred.TabIndex = 49;
			this.rdgContadoCred.Validated += new System.EventHandler(this.rdgContadoCred_Validated);
			// 
			// btPago
			// 
			this.btPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPago.Location = new System.Drawing.Point(192, 16);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(48, 23);
			this.btPago.TabIndex = 46;
			this.btPago.Text = "Pagos";
			this.btPago.Click += new System.EventHandler(this.btPagos_Click);
			this.btPago.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btPago_MouseUp);
			// 
			// btFormaPago
			// 
			this.btFormaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFormaPago.Location = new System.Drawing.Point(192, 48);
			this.btFormaPago.Name = "btFormaPago";
			this.btFormaPago.Size = new System.Drawing.Size(24, 23);
			this.btFormaPago.TabIndex = 45;
			this.btFormaPago.Text = "...";
			this.btFormaPago.Click += new System.EventHandler(this.btFormaPago_Click);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.label5);
			this.ultraTabPageControl3.Controls.Add(this.label4);
			this.ultraTabPageControl3.Controls.Add(this.btCorresponsal);
			this.ultraTabPageControl3.Controls.Add(this.btComision);
			this.ultraTabPageControl3.Controls.Add(this.btDescuento);
			this.ultraTabPageControl3.Controls.Add(this.btFlete);
			this.ultraTabPageControl3.Controls.Add(this.txtDescPorcIva);
			this.ultraTabPageControl3.Controls.Add(this.txtDescPorcIva0);
			this.ultraTabPageControl3.Controls.Add(this.txtDescLinea0);
			this.ultraTabPageControl3.Controls.Add(this.txtDescLinea);
			this.ultraTabPageControl3.Controls.Add(this.txtDescIva);
			this.ultraTabPageControl3.Controls.Add(this.txtDescIva0);
			this.ultraTabPageControl3.Controls.Add(this.txtFlete);
			this.ultraTabPageControl3.Controls.Add(this.btRefresaI);
			this.ultraTabPageControl3.Controls.Add(this.btRefrescaD);
			this.ultraTabPageControl3.Controls.Add(this.label17);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(486, 115);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(192, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(31, 16);
			this.label5.TabIndex = 119;
			this.label5.Text = "Valor";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(304, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(14, 16);
			this.label4.TabIndex = 118;
			this.label4.Text = "%";
			// 
			// btCorresponsal
			// 
			this.btCorresponsal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btCorresponsal.Location = new System.Drawing.Point(328, 80);
			this.btCorresponsal.Name = "btCorresponsal";
			this.btCorresponsal.Size = new System.Drawing.Size(88, 23);
			this.btCorresponsal.TabIndex = 117;
			this.btCorresponsal.Text = "Corresponsal";
			this.btCorresponsal.Click += new System.EventHandler(this.btCorresponsal_Click);
			// 
			// btComision
			// 
			this.btComision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btComision.Location = new System.Drawing.Point(232, 80);
			this.btComision.Name = "btComision";
			this.btComision.Size = new System.Drawing.Size(88, 23);
			this.btComision.TabIndex = 116;
			this.btComision.Text = "Comisión";
			this.btComision.Click += new System.EventHandler(this.btComision_Click);
			// 
			// btDescuento
			// 
			this.btDescuento.Location = new System.Drawing.Point(16, 24);
			this.btDescuento.Name = "btDescuento";
			this.btDescuento.Size = new System.Drawing.Size(88, 23);
			this.btDescuento.TabIndex = 115;
			this.btDescuento.Text = "Descuento Iva";
			this.btDescuento.Click += new System.EventHandler(this.btDescuento_Click);
			// 
			// btFlete
			// 
			this.btFlete.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btFlete.Location = new System.Drawing.Point(16, 80);
			this.btFlete.Name = "btFlete";
			this.btFlete.Size = new System.Drawing.Size(80, 23);
			this.btFlete.TabIndex = 114;
			this.btFlete.Text = "Flete";
			this.btFlete.Click += new System.EventHandler(this.btFlete_Click);
			this.btFlete.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btFlete_MouseDown);
			// 
			// txtDescPorcIva
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescPorcIva.Appearance = appearance14;
			this.txtDescPorcIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DescPorcIva"));
			this.txtDescPorcIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescPorcIva.FormatString = "#,##0.00";
			this.txtDescPorcIva.Location = new System.Drawing.Point(272, 32);
			this.txtDescPorcIva.Name = "txtDescPorcIva";
			this.txtDescPorcIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescPorcIva.PromptChar = ' ';
			this.txtDescPorcIva.Size = new System.Drawing.Size(48, 21);
			this.txtDescPorcIva.TabIndex = 113;
			// 
			// txtDescPorcIva0
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescPorcIva0.Appearance = appearance15;
			this.txtDescPorcIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DescPorcIva0"));
			this.txtDescPorcIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescPorcIva0.FormatString = "#,##0.00";
			this.txtDescPorcIva0.Location = new System.Drawing.Point(272, 56);
			this.txtDescPorcIva0.Name = "txtDescPorcIva0";
			this.txtDescPorcIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescPorcIva0.PromptChar = ' ';
			this.txtDescPorcIva0.Size = new System.Drawing.Size(48, 21);
			this.txtDescPorcIva0.TabIndex = 112;
			// 
			// txtDescLinea0
			// 
			this.txtDescLinea0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc4"));
			this.txtDescLinea0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescLinea0.FormatString = "#,##0.00";
			this.txtDescLinea0.Location = new System.Drawing.Point(328, 56);
			this.txtDescLinea0.Name = "txtDescLinea0";
			this.txtDescLinea0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescLinea0.PromptChar = ' ';
			this.txtDescLinea0.ReadOnly = true;
			this.txtDescLinea0.Size = new System.Drawing.Size(80, 21);
			this.txtDescLinea0.TabIndex = 111;
			// 
			// txtDescLinea
			// 
			this.txtDescLinea.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc3"));
			this.txtDescLinea.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescLinea.FormatString = "#,##0.00";
			this.txtDescLinea.Location = new System.Drawing.Point(328, 32);
			this.txtDescLinea.Name = "txtDescLinea";
			this.txtDescLinea.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescLinea.PromptChar = ' ';
			this.txtDescLinea.ReadOnly = true;
			this.txtDescLinea.Size = new System.Drawing.Size(80, 21);
			this.txtDescLinea.TabIndex = 110;
			// 
			// txtDescIva
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva.Appearance = appearance16;
			this.txtDescIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc1"));
			this.txtDescIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva.FormatString = "#,##0.00";
			this.txtDescIva.Location = new System.Drawing.Point(128, 32);
			this.txtDescIva.Name = "txtDescIva";
			this.txtDescIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva.PromptChar = ' ';
			this.txtDescIva.Size = new System.Drawing.Size(96, 21);
			this.txtDescIva.TabIndex = 109;
			this.txtDescIva.Validated += new System.EventHandler(this.btRefrescaD_Click);
			// 
			// txtDescIva0
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0.Appearance = appearance17;
			this.txtDescIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc2"));
			this.txtDescIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0.FormatString = "#,##0.00";
			this.txtDescIva0.Location = new System.Drawing.Point(128, 56);
			this.txtDescIva0.Name = "txtDescIva0";
			this.txtDescIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0.PromptChar = ' ';
			this.txtDescIva0.Size = new System.Drawing.Size(96, 21);
			this.txtDescIva0.TabIndex = 108;
			// 
			// txtFlete
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFlete.Appearance = appearance18;
			this.txtFlete.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Flete"));
			this.txtFlete.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFlete.FormatString = "#,##0.00";
			this.txtFlete.Location = new System.Drawing.Point(128, 80);
			this.txtFlete.Name = "txtFlete";
			this.txtFlete.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFlete.PromptChar = ' ';
			this.txtFlete.Size = new System.Drawing.Size(96, 21);
			this.txtFlete.TabIndex = 107;
			// 
			// btRefresaI
			// 
			this.btRefresaI.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRefresaI.Location = new System.Drawing.Point(231, 56);
			this.btRefresaI.Name = "btRefresaI";
			this.btRefresaI.Size = new System.Drawing.Size(32, 23);
			this.btRefresaI.TabIndex = 81;
			this.btRefresaI.Text = "<-";
			this.btRefresaI.Click += new System.EventHandler(this.btRefrescaI_Click);
			// 
			// btRefrescaD
			// 
			this.btRefrescaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRefrescaD.Location = new System.Drawing.Point(231, 32);
			this.btRefrescaD.Name = "btRefrescaD";
			this.btRefrescaD.Size = new System.Drawing.Size(32, 23);
			this.btRefrescaD.TabIndex = 80;
			this.btRefrescaD.Text = "->";
			this.btRefrescaD.Click += new System.EventHandler(this.btRefrescaD_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(16, 56);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(89, 16);
			this.label17.TabIndex = 79;
			this.label17.Text = "Descuento IVA 0";
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.cmbRecibeFabara);
			this.pgEntrega.Controls.Add(this.btComprob1);
			this.pgEntrega.Controls.Add(this.txtComprob1);
			this.pgEntrega.Controls.Add(this.btImportacion);
			this.pgEntrega.Controls.Add(this.label38);
			this.pgEntrega.Controls.Add(this.cmbRecibe);
			this.pgEntrega.Controls.Add(this.label37);
			this.pgEntrega.Controls.Add(this.txtOtro);
			this.pgEntrega.Controls.Add(this.txtDepartamento);
			this.pgEntrega.Controls.Add(this.chkRevisado);
			this.pgEntrega.Controls.Add(this.chkEntregado);
			this.pgEntrega.Controls.Add(this.dtFechaRevision);
			this.pgEntrega.Controls.Add(this.dtFechaEntrega);
			this.pgEntrega.Controls.Add(this.label20);
			this.pgEntrega.Controls.Add(this.label19);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(486, 115);
			// 
			// cmbRecibeFabara
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRecibeFabara.Appearance = appearance19;
			this.cmbRecibeFabara.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecibeFabara.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRecibeFabara.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idRecibe"));
			this.cmbRecibeFabara.DisplayMember = "";
			this.cmbRecibeFabara.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecibeFabara.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecibeFabara.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecibeFabara.Location = new System.Drawing.Point(136, 88);
			this.cmbRecibeFabara.Name = "cmbRecibeFabara";
			this.cmbRecibeFabara.Size = new System.Drawing.Size(304, 21);
			this.cmbRecibeFabara.TabIndex = 76;
			this.cmbRecibeFabara.ValueMember = "";
			this.cmbRecibeFabara.Visible = false;
			this.cmbRecibeFabara.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbRecibeFabara_MouseUp);
			// 
			// btComprob1
			// 
			this.btComprob1.Location = new System.Drawing.Point(344, 8);
			this.btComprob1.Name = "btComprob1";
			this.btComprob1.Size = new System.Drawing.Size(80, 23);
			this.btComprob1.TabIndex = 75;
			this.btComprob1.Text = "Comprob. 1";
			this.btComprob1.Click += new System.EventHandler(this.btComprob1_Click);
			// 
			// txtComprob1
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprob1.Appearance = appearance20;
			this.txtComprob1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprob1"));
			this.txtComprob1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprob1.Location = new System.Drawing.Point(328, 32);
			this.txtComprob1.Name = "txtComprob1";
			this.txtComprob1.Size = new System.Drawing.Size(104, 21);
			this.txtComprob1.TabIndex = 74;
			// 
			// btImportacion
			// 
			this.btImportacion.Location = new System.Drawing.Point(16, 64);
			this.btImportacion.Name = "btImportacion";
			this.btImportacion.Size = new System.Drawing.Size(80, 23);
			this.btImportacion.TabIndex = 73;
			this.btImportacion.Text = "Importación";
			this.btImportacion.Click += new System.EventHandler(this.btImportacion_Click);
			this.btImportacion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btImportacion_MouseUp);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(24, 88);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(43, 16);
			this.label38.TabIndex = 72;
			this.label38.Text = "Recibe:";
			// 
			// cmbRecibe
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRecibe.Appearance = appearance21;
			this.cmbRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecibe.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRecibe.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idRecibe"));
			this.cmbRecibe.DataSource = this.cdvEntrega;
			ultraGridBand4.AddButtonCaption = "Cliente";
			this.cmbRecibe.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRecibe.DisplayMember = "Nombre";
			this.cmbRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecibe.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecibe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecibe.Location = new System.Drawing.Point(112, 88);
			this.cmbRecibe.Name = "cmbRecibe";
			this.cmbRecibe.Size = new System.Drawing.Size(304, 21);
			this.cmbRecibe.TabIndex = 71;
			this.cmbRecibe.ValueMember = "idCliente";
			this.cmbRecibe.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbRecibe_MouseUp);
			// 
			// cdvEntrega
			// 
			this.cdvEntrega.BindingContextCtrl = this;
			this.cdvEntrega.DataSet = null;
			this.cdvEntrega.RowFilter = "Entrega = 1";
			this.cdvEntrega.Sort = "Nombre";
			this.cdvEntrega.TableName = "";
			this.cdvEntrega.TableViewName = "Cliente";
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(248, 64);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(26, 16);
			this.label37.TabIndex = 58;
			this.label37.Text = "Otro";
			// 
			// txtOtro
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtro.Appearance = appearance22;
			this.txtOtro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Otro"));
			this.txtOtro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtro.Location = new System.Drawing.Point(288, 64);
			this.txtOtro.Name = "txtOtro";
			this.txtOtro.Size = new System.Drawing.Size(128, 21);
			this.txtOtro.TabIndex = 57;
			// 
			// txtDepartamento
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepartamento.Appearance = appearance23;
			this.txtDepartamento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Departamento"));
			this.txtDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepartamento.Location = new System.Drawing.Point(112, 64);
			this.txtDepartamento.Name = "txtDepartamento";
			this.txtDepartamento.Size = new System.Drawing.Size(128, 21);
			this.txtDepartamento.TabIndex = 55;
			// 
			// chkRevisado
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRevisado.Appearance = appearance24;
			this.chkRevisado.BackColor = System.Drawing.Color.Transparent;
			this.chkRevisado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Revisado"));
			this.chkRevisado.Location = new System.Drawing.Point(232, 40);
			this.chkRevisado.Name = "chkRevisado";
			this.chkRevisado.Size = new System.Drawing.Size(80, 20);
			this.chkRevisado.TabIndex = 53;
			this.chkRevisado.Text = "Revisado";
			// 
			// chkEntregado
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance25;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Entregado"));
			this.chkEntregado.Location = new System.Drawing.Point(232, 16);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(80, 20);
			this.chkEntregado.TabIndex = 52;
			this.chkEntregado.Text = "Entregado";
			// 
			// dtFechaRevision
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRevision.Appearance = appearance26;
			this.dtFechaRevision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaRevision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaRevision"));
			dateButton1.Caption = "Today";
			this.dtFechaRevision.DateButtons.Add(dateButton1);
			this.dtFechaRevision.Format = "dd/MMM/yyyy";
			this.dtFechaRevision.Location = new System.Drawing.Point(112, 40);
			this.dtFechaRevision.Name = "dtFechaRevision";
			this.dtFechaRevision.NonAutoSizeHeight = 23;
			this.dtFechaRevision.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRevision.SpinButtonsVisible = true;
			this.dtFechaRevision.TabIndex = 51;
			this.dtFechaRevision.Value = new System.DateTime(2005, 7, 1, 0, 0, 0, 0);
			// 
			// dtFechaEntrega
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance27;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaEntrega"));
			dateButton2.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton2);
			this.dtFechaEntrega.Format = "dd/MMM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(112, 16);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 50;
			this.dtFechaEntrega.Value = new System.DateTime(2005, 7, 1, 0, 0, 0, 0);
			// 
			// label20
			// 
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(16, 40);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(98, 16);
			this.label20.TabIndex = 45;
			this.label20.Text = "Fecha de Revisión";
			// 
			// label19
			// 
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(16, 16);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(94, 16);
			this.label19.TabIndex = 44;
			this.label19.Text = "Fecha de Entrega";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor1);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor2);
			this.ultraTabPageControl5.Controls.Add(this.label2);
			this.ultraTabPageControl5.Controls.Add(this.label3);
			this.ultraTabPageControl5.Controls.Add(this.txtInteresPagar);
			this.ultraTabPageControl5.Controls.Add(this.txtEmpezarEn);
			this.ultraTabPageControl5.Controls.Add(this.txtPagosCada);
			this.ultraTabPageControl5.Controls.Add(this.txtPorcInteres);
			this.ultraTabPageControl5.Controls.Add(this.txtTotalCuotas);
			this.ultraTabPageControl5.Controls.Add(this.txtValorCuotas);
			this.ultraTabPageControl5.Controls.Add(this.txtAnticipo);
			this.ultraTabPageControl5.Controls.Add(this.txtNumCuotas);
			this.ultraTabPageControl5.Controls.Add(this.btTotalCuotas);
			this.ultraTabPageControl5.Controls.Add(this.label25);
			this.ultraTabPageControl5.Controls.Add(this.label26);
			this.ultraTabPageControl5.Controls.Add(this.label27);
			this.ultraTabPageControl5.Controls.Add(this.label28);
			this.ultraTabPageControl5.Controls.Add(this.label23);
			this.ultraTabPageControl5.Controls.Add(this.label22);
			this.ultraTabPageControl5.Controls.Add(this.label21);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(486, 115);
			// 
			// ultraNumericEditor1
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance28;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Interes2"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(432, 56);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor1.TabIndex = 121;
			// 
			// ultraNumericEditor2
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance29;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Dias2"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(312, 56);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor2.TabIndex = 120;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(376, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 119;
			this.label2.Text = "Interes 2";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(224, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 118;
			this.label3.Text = "Dìas 2:";
			// 
			// txtInteresPagar
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInteresPagar.Appearance = appearance30;
			this.txtInteresPagar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ValorInteres"));
			this.txtInteresPagar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteresPagar.FormatString = "#,##0.00";
			this.txtInteresPagar.Location = new System.Drawing.Point(312, 80);
			this.txtInteresPagar.Name = "txtInteresPagar";
			this.txtInteresPagar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteresPagar.PromptChar = ' ';
			this.txtInteresPagar.Size = new System.Drawing.Size(96, 21);
			this.txtInteresPagar.TabIndex = 117;
			// 
			// txtEmpezarEn
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpezarEn.Appearance = appearance31;
			this.txtEmpezarEn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DiasInicio"));
			this.txtEmpezarEn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpezarEn.FormatString = "#,##0.00";
			this.txtEmpezarEn.Location = new System.Drawing.Point(432, 32);
			this.txtEmpezarEn.Name = "txtEmpezarEn";
			this.txtEmpezarEn.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEmpezarEn.PromptChar = ' ';
			this.txtEmpezarEn.Size = new System.Drawing.Size(48, 21);
			this.txtEmpezarEn.TabIndex = 116;
			// 
			// txtPagosCada
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPagosCada.Appearance = appearance32;
			this.txtPagosCada.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.PagosCada"));
			this.txtPagosCada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagosCada.FormatString = "#,##0.00";
			this.txtPagosCada.Location = new System.Drawing.Point(312, 32);
			this.txtPagosCada.Name = "txtPagosCada";
			this.txtPagosCada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPagosCada.PromptChar = ' ';
			this.txtPagosCada.Size = new System.Drawing.Size(48, 21);
			this.txtPagosCada.TabIndex = 115;
			// 
			// txtPorcInteres
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcInteres.Appearance = appearance33;
			this.txtPorcInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Interes"));
			this.txtPorcInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcInteres.FormatString = "#,##0.00";
			this.txtPorcInteres.Location = new System.Drawing.Point(312, 8);
			this.txtPorcInteres.Name = "txtPorcInteres";
			this.txtPorcInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcInteres.PromptChar = ' ';
			this.txtPorcInteres.Size = new System.Drawing.Size(96, 21);
			this.txtPorcInteres.TabIndex = 114;
			// 
			// txtTotalCuotas
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalCuotas.Appearance = appearance34;
			this.txtTotalCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.TotalCuotas"));
			this.txtTotalCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalCuotas.FormatString = "#,##0.00";
			this.txtTotalCuotas.Location = new System.Drawing.Point(104, 80);
			this.txtTotalCuotas.Name = "txtTotalCuotas";
			this.txtTotalCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalCuotas.PromptChar = ' ';
			this.txtTotalCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtTotalCuotas.TabIndex = 113;
			// 
			// txtValorCuotas
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCuotas.Appearance = appearance35;
			this.txtValorCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ValorCuotas"));
			this.txtValorCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCuotas.FormatString = "#,##0.00";
			this.txtValorCuotas.Location = new System.Drawing.Point(104, 56);
			this.txtValorCuotas.Name = "txtValorCuotas";
			this.txtValorCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCuotas.PromptChar = ' ';
			this.txtValorCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtValorCuotas.TabIndex = 112;
			// 
			// txtAnticipo
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnticipo.Appearance = appearance36;
			this.txtAnticipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Contado"));
			this.txtAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnticipo.FormatString = "#,##0.00";
			this.txtAnticipo.Location = new System.Drawing.Point(104, 32);
			this.txtAnticipo.Name = "txtAnticipo";
			this.txtAnticipo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAnticipo.PromptChar = ' ';
			this.txtAnticipo.Size = new System.Drawing.Size(96, 21);
			this.txtAnticipo.TabIndex = 111;
			// 
			// txtNumCuotas
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance37;
			this.txtNumCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCuotas"));
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.FormatString = "#,##0.00";
			this.txtNumCuotas.Location = new System.Drawing.Point(104, 8);
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtNumCuotas.TabIndex = 110;
			// 
			// btTotalCuotas
			// 
			this.btTotalCuotas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTotalCuotas.Location = new System.Drawing.Point(8, 80);
			this.btTotalCuotas.Name = "btTotalCuotas";
			this.btTotalCuotas.Size = new System.Drawing.Size(73, 23);
			this.btTotalCuotas.TabIndex = 106;
			this.btTotalCuotas.Text = "Total Pagar";
			this.btTotalCuotas.Click += new System.EventHandler(this.btTotalCuotas_Click);
			this.btTotalCuotas.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTotalCuotas_MouseDown);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(223, 82);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(81, 16);
			this.label25.TabIndex = 104;
			this.label25.Text = "Interes a Pagar";
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(376, 32);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(49, 16);
			this.label26.TabIndex = 102;
			this.label26.Text = "Empezar";
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(223, 34);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(70, 16);
			this.label27.TabIndex = 100;
			this.label27.Text = "Pagos Cada:";
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(223, 10);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(67, 16);
			this.label28.TabIndex = 98;
			this.label28.Text = "% de Interes";
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(15, 58);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(70, 16);
			this.label23.TabIndex = 95;
			this.label23.Text = "Valor Cuotas";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(15, 34);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(45, 16);
			this.label22.TabIndex = 93;
			this.label22.Text = "Anticipo";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(15, 10);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(61, 16);
			this.label21.TabIndex = 91;
			this.label21.Text = "No. Cuotas";
			// 
			// TabPageCOA
			// 
			this.TabPageCOA.Controls.Add(this.label40);
			this.TabPageCOA.Controls.Add(this.cmbIce);
			this.TabPageCOA.Controls.Add(this.chkConvenioInt);
			this.TabPageCOA.Controls.Add(this.chkDevolucion);
			this.TabPageCOA.Controls.Add(this.OptBienServicio);
			this.TabPageCOA.Controls.Add(this.txtRefrendo);
			this.TabPageCOA.Controls.Add(this.label30);
			this.TabPageCOA.Controls.Add(this.label29);
			this.TabPageCOA.Controls.Add(this.cmbCredTribut);
			this.TabPageCOA.Location = new System.Drawing.Point(-10000, -10000);
			this.TabPageCOA.Name = "TabPageCOA";
			this.TabPageCOA.Size = new System.Drawing.Size(486, 115);
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(232, 8);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(23, 16);
			this.label40.TabIndex = 95;
			this.label40.Text = "ICE";
			// 
			// cmbIce
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbIce.Appearance = appearance38;
			this.cmbIce.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbIce.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbIce.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idPorcIce"));
			this.cmbIce.DataMember = "PorcIce";
			this.cmbIce.DataSource = this.cdsCompraTabla;
			ultraGridBand5.AddButtonCaption = "CompraCredTribut";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbIce.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbIce.DisplayMember = "concepto";
			this.cmbIce.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbIce.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbIce.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbIce.Location = new System.Drawing.Point(264, 8);
			this.cmbIce.Name = "cmbIce";
			this.cmbIce.Size = new System.Drawing.Size(144, 21);
			this.cmbIce.TabIndex = 94;
			this.cmbIce.ValueMember = "idPorcIce";
			this.cmbIce.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbIce_MouseDown);
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
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsComprobante_BeforeFill);
			// 
			// chkConvenioInt
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConvenioInt.Appearance = appearance39;
			this.chkConvenioInt.BackColor = System.Drawing.Color.Transparent;
			this.chkConvenioInt.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.ConvenioInternac"));
			this.chkConvenioInt.Location = new System.Drawing.Point(232, 64);
			this.chkConvenioInt.Name = "chkConvenioInt";
			this.chkConvenioInt.Size = new System.Drawing.Size(160, 20);
			this.chkConvenioInt.TabIndex = 93;
			this.chkConvenioInt.Text = "Convenio Internacional";
			// 
			// chkDevolucion
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDevolucion.Appearance = appearance40;
			this.chkDevolucion.BackColor = System.Drawing.Color.Transparent;
			this.chkDevolucion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.DerechoDevolucion"));
			this.chkDevolucion.Location = new System.Drawing.Point(232, 40);
			this.chkDevolucion.Name = "chkDevolucion";
			this.chkDevolucion.Size = new System.Drawing.Size(160, 20);
			this.chkDevolucion.TabIndex = 92;
			this.chkDevolucion.Text = "Derecho a Devolución";
			// 
			// OptBienServicio
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.OptBienServicio.Appearance = appearance41;
			this.OptBienServicio.BackColor = System.Drawing.Color.Transparent;
			this.OptBienServicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.BienServicio"));
			this.OptBienServicio.ItemAppearance = appearance42;
			this.OptBienServicio.ItemOrigin = new System.Drawing.Point(4, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Transferencia de Bienes";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Prestación de Servicios";
			this.OptBienServicio.Items.Add(valueListItem3);
			this.OptBienServicio.Items.Add(valueListItem4);
			this.OptBienServicio.ItemSpacingVertical = 6;
			this.OptBienServicio.Location = new System.Drawing.Point(16, 40);
			this.OptBienServicio.Name = "OptBienServicio";
			this.OptBienServicio.Size = new System.Drawing.Size(160, 48);
			this.OptBienServicio.TabIndex = 91;
			// 
			// txtRefrendo
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRefrendo.Appearance = appearance43;
			this.txtRefrendo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Refrendo"));
			this.txtRefrendo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRefrendo.Location = new System.Drawing.Point(288, 88);
			this.txtRefrendo.Name = "txtRefrendo";
			this.txtRefrendo.Size = new System.Drawing.Size(120, 21);
			this.txtRefrendo.TabIndex = 90;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(232, 88);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(51, 16);
			this.label30.TabIndex = 53;
			this.label30.Text = "Refrendo";
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(8, 8);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(58, 16);
			this.label29.TabIndex = 52;
			this.label29.Text = "Créd. Trib.";
			// 
			// cmbCredTribut
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCredTribut.Appearance = appearance44;
			this.cmbCredTribut.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCredTribut.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCredTributario"));
			this.cmbCredTribut.DataMember = "CompraCredTribut";
			this.cmbCredTribut.DataSource = this.cdsCompraTabla;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 55;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 24;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Width = 329;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbCredTribut.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCredTribut.DisplayMember = "CredTributario";
			this.cmbCredTribut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribut.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribut.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTribut.Location = new System.Drawing.Point(72, 8);
			this.cmbCredTribut.Name = "cmbCredTribut";
			this.cmbCredTribut.Size = new System.Drawing.Size(144, 21);
			this.cmbCredTribut.TabIndex = 51;
			this.cmbCredTribut.ValueMember = "idCredTributario";
			this.cmbCredTribut.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCredTribut_MouseDown);
			// 
			// tbpAutorizacion
			// 
			this.tbpAutorizacion.Controls.Add(this.txtUsuario);
			this.tbpAutorizacion.Controls.Add(this.label39);
			this.tbpAutorizacion.Controls.Add(this.label11);
			this.tbpAutorizacion.Controls.Add(this.cmbCaducidad);
			this.tbpAutorizacion.Controls.Add(this.chkAutPresupuestaria);
			this.tbpAutorizacion.Controls.Add(this.chkAutPrincipal);
			this.tbpAutorizacion.Controls.Add(this.spnPrioridad);
			this.tbpAutorizacion.Controls.Add(this.label34);
			this.tbpAutorizacion.Controls.Add(this.txtSerieFact);
			this.tbpAutorizacion.Controls.Add(this.label33);
			this.tbpAutorizacion.Controls.Add(this.txtAutImprenta);
			this.tbpAutorizacion.Controls.Add(this.label32);
			this.tbpAutorizacion.Controls.Add(this.txtAutFactura);
			this.tbpAutorizacion.Controls.Add(this.label31);
			this.tbpAutorizacion.Location = new System.Drawing.Point(-10000, -10000);
			this.tbpAutorizacion.Name = "tbpAutorizacion";
			this.tbpAutorizacion.Size = new System.Drawing.Size(486, 115);
			// 
			// txtUsuario
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance45;
			this.txtUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Usuario"));
			this.txtUsuario.Location = new System.Drawing.Point(288, 80);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.ReadOnly = true;
			this.txtUsuario.Size = new System.Drawing.Size(136, 21);
			this.txtUsuario.TabIndex = 103;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(232, 80);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(43, 16);
			this.label39.TabIndex = 102;
			this.label39.Text = "Usuario";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(16, 80);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(71, 16);
			this.label11.TabIndex = 101;
			this.label11.Text = "F. Caducidad";
			// 
			// cmbCaducidad
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCaducidad.Appearance = appearance46;
			this.cmbCaducidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaCaducidad"));
			this.cmbCaducidad.DateTime = new System.DateTime(2006, 3, 17, 0, 0, 0, 0);
			this.cmbCaducidad.Location = new System.Drawing.Point(96, 80);
			this.cmbCaducidad.Name = "cmbCaducidad";
			this.cmbCaducidad.Size = new System.Drawing.Size(120, 21);
			this.cmbCaducidad.TabIndex = 100;
			this.cmbCaducidad.Value = new System.DateTime(2006, 3, 17, 0, 0, 0, 0);
			// 
			// chkAutPresupuestaria
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutPresupuestaria.Appearance = appearance47;
			this.chkAutPresupuestaria.BackColor = System.Drawing.Color.Transparent;
			this.chkAutPresupuestaria.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.AutPresupuesto"));
			this.chkAutPresupuestaria.Location = new System.Drawing.Point(232, 56);
			this.chkAutPresupuestaria.Name = "chkAutPresupuestaria";
			this.chkAutPresupuestaria.Size = new System.Drawing.Size(176, 20);
			this.chkAutPresupuestaria.TabIndex = 99;
			this.chkAutPresupuestaria.Text = "Autorización Presupuestaria";
			// 
			// chkAutPrincipal
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutPrincipal.Appearance = appearance48;
			this.chkAutPrincipal.BackColor = System.Drawing.Color.Transparent;
			this.chkAutPrincipal.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.AutPrincipal"));
			this.chkAutPrincipal.Location = new System.Drawing.Point(232, 32);
			this.chkAutPrincipal.Name = "chkAutPrincipal";
			this.chkAutPrincipal.Size = new System.Drawing.Size(160, 20);
			this.chkAutPrincipal.TabIndex = 98;
			this.chkAutPrincipal.Text = "Autorización Principal";
			this.chkAutPrincipal.Click += new System.EventHandler(this.chkAutPrincipal_Click);
			// 
			// spnPrioridad
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnPrioridad.Appearance = appearance49;
			this.spnPrioridad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Prioridad"));
			this.spnPrioridad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnPrioridad.Location = new System.Drawing.Point(304, 8);
			this.spnPrioridad.MaskInput = "nn";
			this.spnPrioridad.Name = "spnPrioridad";
			this.spnPrioridad.PromptChar = ' ';
			this.spnPrioridad.Size = new System.Drawing.Size(56, 21);
			this.spnPrioridad.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPrioridad.TabIndex = 97;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(232, 8);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(50, 16);
			this.label34.TabIndex = 96;
			this.label34.Text = "Prioridad";
			// 
			// txtSerieFact
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance50;
			this.txtSerieFact.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SerieFactura"));
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Location = new System.Drawing.Point(96, 56);
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(120, 21);
			this.txtSerieFact.TabIndex = 95;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(16, 56);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(72, 16);
			this.label33.TabIndex = 94;
			this.label33.Text = "Serie Factura";
			// 
			// txtAutImprenta
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutImprenta.Appearance = appearance51;
			this.txtAutImprenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.AutImprenta"));
			this.txtAutImprenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutImprenta.Location = new System.Drawing.Point(96, 32);
			this.txtAutImprenta.Name = "txtAutImprenta";
			this.txtAutImprenta.Size = new System.Drawing.Size(120, 21);
			this.txtAutImprenta.TabIndex = 93;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(16, 32);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(72, 16);
			this.label32.TabIndex = 92;
			this.label32.Text = "Aut. Imprenta";
			// 
			// txtAutFactura
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance52;
			this.txtAutFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.AutFactura"));
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(96, 8);
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(120, 21);
			this.txtAutFactura.TabIndex = 91;
			this.txtAutFactura.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtAutFactura_MouseDown);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(16, 8);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(66, 16);
			this.label31.TabIndex = 53;
			this.label31.Text = "Aut. Factura";
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.chkCodBarras);
			this.ultraTabPageControl6.Controls.Add(this.btIngresar);
			this.ultraTabPageControl6.Controls.Add(this.btBuscar);
			this.ultraTabPageControl6.Controls.Add(this.txtBusca);
			this.ultraTabPageControl6.Controls.Add(this.ugrBusca);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(486, 115);
			// 
			// chkCodBarras
			// 
			this.chkCodBarras.Location = new System.Drawing.Point(8, 88);
			this.chkCodBarras.Name = "chkCodBarras";
			this.chkCodBarras.Size = new System.Drawing.Size(88, 20);
			this.chkCodBarras.TabIndex = 4;
			this.chkCodBarras.Text = "Cod. Barras";
			// 
			// btIngresar
			// 
			this.btIngresar.Location = new System.Drawing.Point(16, 64);
			this.btIngresar.Name = "btIngresar";
			this.btIngresar.Size = new System.Drawing.Size(72, 23);
			this.btIngresar.TabIndex = 3;
			this.btIngresar.Text = "Ingresar";
			this.btIngresar.Click += new System.EventHandler(this.btIngresar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(16, 32);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 23);
			this.btBuscar.TabIndex = 2;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscarArticulos);
			// 
			// txtBusca
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance53;
			this.txtBusca.Location = new System.Drawing.Point(8, 8);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(88, 21);
			this.txtBusca.TabIndex = 1;
			this.txtBusca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusca_KeyPress);
			// 
			// ugrBusca
			// 
			this.ugrBusca.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrBusca.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrBusca.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			this.ugrBusca.Location = new System.Drawing.Point(104, 8);
			this.ugrBusca.Name = "ugrBusca";
			this.ugrBusca.Size = new System.Drawing.Size(376, 104);
			this.ugrBusca.TabIndex = 0;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.spnNumFrecCopiarCotiz);
			this.ultraTabPageControl4.Controls.Add(this.spnNumCopiarCotiz);
			this.ultraTabPageControl4.Controls.Add(this.btnCopiarCotiz);
			this.ultraTabPageControl4.Controls.Add(this.btnComisionExt);
			this.ultraTabPageControl4.Controls.Add(this.cmbNombre);
			this.ultraTabPageControl4.Controls.Add(this.lblComisionExt);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor6);
			this.ultraTabPageControl4.Controls.Add(this.label46);
			this.ultraTabPageControl4.Controls.Add(this.label45);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor5);
			this.ultraTabPageControl4.Controls.Add(this.label14);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor4);
			this.ultraTabPageControl4.Controls.Add(this.label13);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl4.Controls.Add(this.txtCIF);
			this.ultraTabPageControl4.Controls.Add(this.label44);
			this.ultraTabPageControl4.Controls.Add(this.cmbImportacion);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl4.Controls.Add(this.label43);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl4.Controls.Add(this.label42);
			this.ultraTabPageControl4.Controls.Add(this.txtPeso);
			this.ultraTabPageControl4.Controls.Add(this.label36);
			this.ultraTabPageControl4.Controls.Add(this.label41);
			this.ultraTabPageControl4.Controls.Add(this.txtBulto);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(486, 115);
			// 
			// spnNumFrecCopiarCotiz
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnNumFrecCopiarCotiz.Appearance = appearance54;
			this.spnNumFrecCopiarCotiz.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNumFrecCopiarCotiz.Location = new System.Drawing.Point(272, 88);
			this.spnNumFrecCopiarCotiz.MaskInput = "nn";
			this.spnNumFrecCopiarCotiz.Name = "spnNumFrecCopiarCotiz";
			this.spnNumFrecCopiarCotiz.PromptChar = ' ';
			this.spnNumFrecCopiarCotiz.Size = new System.Drawing.Size(48, 21);
			this.spnNumFrecCopiarCotiz.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumFrecCopiarCotiz.TabIndex = 137;
			this.toolTip.SetToolTip(this.spnNumFrecCopiarCotiz, "Ingresae la frecuencia en numeros MENSUAL, BIMENSUAL, TRIMESTRAL, etc");
			this.spnNumFrecCopiarCotiz.Visible = false;
			// 
			// spnNumCopiarCotiz
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnNumCopiarCotiz.Appearance = appearance55;
			this.spnNumCopiarCotiz.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNumCopiarCotiz.Location = new System.Drawing.Point(328, 88);
			this.spnNumCopiarCotiz.MaskInput = "nn";
			this.spnNumCopiarCotiz.Name = "spnNumCopiarCotiz";
			this.spnNumCopiarCotiz.PromptChar = ' ';
			this.spnNumCopiarCotiz.Size = new System.Drawing.Size(48, 21);
			this.spnNumCopiarCotiz.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumCopiarCotiz.TabIndex = 136;
			this.toolTip.SetToolTip(this.spnNumCopiarCotiz, "INgresae el numero de Meses. de Esto dependera el numero de copias de cotizacion");
			this.spnNumCopiarCotiz.Visible = false;
			// 
			// btnCopiarCotiz
			// 
			this.btnCopiarCotiz.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCopiarCotiz.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btnCopiarCotiz.Location = new System.Drawing.Point(384, 88);
			this.btnCopiarCotiz.Name = "btnCopiarCotiz";
			this.btnCopiarCotiz.Size = new System.Drawing.Size(80, 23);
			this.btnCopiarCotiz.TabIndex = 135;
			this.btnCopiarCotiz.Text = "Copia Cotizacion";
			this.btnCopiarCotiz.Visible = false;
			this.btnCopiarCotiz.Click += new System.EventHandler(this.btnCopiarCotiz_Click);
			// 
			// btnComisionExt
			// 
			this.btnComisionExt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComisionExt.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btnComisionExt.Location = new System.Drawing.Point(384, 88);
			this.btnComisionExt.Name = "btnComisionExt";
			this.btnComisionExt.Size = new System.Drawing.Size(80, 23);
			this.btnComisionExt.TabIndex = 131;
			this.btnComisionExt.Text = "Comis. Exter.";
			this.btnComisionExt.Visible = false;
			this.btnComisionExt.Click += new System.EventHandler(this.btnComisionExt_Click);
			// 
			// cmbNombre
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNombre.Appearance = appearance56;
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idComisionConfiteca"));
			this.cmbNombre.DataSource = this.cdvConfiteca;
			ultraGridBand7.AddButtonCaption = "Cliente";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 231;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(264, 88);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(112, 21);
			this.cmbNombre.TabIndex = 131;
			this.cmbNombre.ValueMember = "idComisionConfiteca";
			this.cmbNombre.Visible = false;
			// 
			// cdvConfiteca
			// 
			this.cdvConfiteca.BindingContextCtrl = this;
			this.cdvConfiteca.DataSet = this.cdsConfiteca;
			this.cdvConfiteca.Sort = "Nombre";
			this.cdvConfiteca.TableName = "";
			this.cdvConfiteca.TableViewName = "P_ConfitecaComision";
			// 
			// cdsConfiteca
			// 
			this.cdsConfiteca.BindingContextCtrl = this;
			this.cdsConfiteca.CaseSensitive = false;
			this.cdsConfiteca.DataLibrary = "LibFacturacion";
			this.cdsConfiteca.DataLibraryUrl = "";
			this.cdsConfiteca.DataSetDef = "dsConfitecaComision";
			this.cdsConfiteca.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsConfiteca.Name = "cdsConfiteca";
			this.cdsConfiteca.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsConfiteca.SchemaDef = null;
			// 
			// lblComisionExt
			// 
			this.lblComisionExt.AutoSize = true;
			this.lblComisionExt.BackColor = System.Drawing.Color.Transparent;
			this.lblComisionExt.Location = new System.Drawing.Point(176, 88);
			this.lblComisionExt.Name = "lblComisionExt";
			this.lblComisionExt.Size = new System.Drawing.Size(84, 16);
			this.lblComisionExt.TabIndex = 80;
			this.lblComisionExt.Text = "Comisión Exter.";
			this.lblComisionExt.Visible = false;
			// 
			// ultraTextEditor6
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor6.Appearance = appearance57;
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje2"));
			this.ultraTextEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor6.Location = new System.Drawing.Point(360, 16);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(88, 21);
			this.ultraTextEditor6.TabIndex = 7;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(328, 16);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(31, 16);
			this.label46.TabIndex = 79;
			this.label46.Text = "Emb.";
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(184, 64);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(21, 16);
			this.label45.TabIndex = 78;
			this.label45.Text = "Via";
			// 
			// ultraTextEditor5
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor5.Appearance = appearance58;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje3"));
			this.ultraTextEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor5.Location = new System.Drawing.Point(224, 64);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor5.TabIndex = 6;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(184, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(31, 16);
			this.label14.TabIndex = 76;
			this.label14.Text = "AWG";
			// 
			// ultraTextEditor4
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor4.Appearance = appearance59;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje1"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(224, 40);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor4.TabIndex = 5;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(184, 16);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(28, 16);
			this.label13.TabIndex = 72;
			this.label13.Text = "Guía";
			// 
			// ultraTextEditor2
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor2.Appearance = appearance60;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Otro"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(224, 16);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor2.TabIndex = 4;
			// 
			// txtCIF
			// 
			this.txtCIF.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Cif"));
			this.txtCIF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCIF.FormatString = "#,##0.00";
			this.txtCIF.Location = new System.Drawing.Point(360, 64);
			this.txtCIF.Name = "txtCIF";
			this.txtCIF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCIF.PromptChar = ' ';
			this.txtCIF.Size = new System.Drawing.Size(88, 21);
			this.txtCIF.TabIndex = 8;
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(328, 64);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(22, 16);
			this.label44.TabIndex = 69;
			this.label44.Text = "CIF";
			// 
			// cmbImportacion
			// 
			this.cmbImportacion.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbImportacion.DataMember = "ImportacionTipo";
			this.cmbImportacion.DataSource = this.cdsTipoImportacion;
			this.cmbImportacion.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Width = 278;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Hidden = true;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbImportacion.DisplayMember = "Nombre";
			this.cmbImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbImportacion.Location = new System.Drawing.Point(168, 16);
			this.cmbImportacion.Name = "cmbImportacion";
			this.cmbImportacion.Size = new System.Drawing.Size(280, 96);
			this.cmbImportacion.TabIndex = 68;
			this.cmbImportacion.ValueMember = "idImportacionTipo";
			this.cmbImportacion.Visible = false;
			// 
			// cdsTipoImportacion
			// 
			this.cdsTipoImportacion.BindingContextCtrl = this;
			this.cdsTipoImportacion.DataLibrary = "LibFacturacion";
			this.cdsTipoImportacion.DataLibraryUrl = "";
			this.cdsTipoImportacion.DataSetDef = "dsTipoImportacion";
			this.cdsTipoImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoImportacion.Name = "cdsTipoImportacion";
			this.cdsTipoImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoImportacion.SchemaDef = null;
			this.cdsTipoImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoImportacion_BeforeFill);
			// 
			// ultraTextEditor3
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor3.Appearance = appearance61;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Depart"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(72, 88);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(100, 21);
			this.ultraTextEditor3.TabIndex = 3;
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Location = new System.Drawing.Point(16, 88);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(41, 16);
			this.label43.TabIndex = 62;
			this.label43.Text = "Depart.";
			// 
			// ultraTextEditor1
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance62;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Pedido"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(72, 64);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(100, 21);
			this.ultraTextEditor1.TabIndex = 2;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(16, 64);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(40, 16);
			this.label42.TabIndex = 60;
			this.label42.Text = "Pedido";
			// 
			// txtPeso
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPeso.Appearance = appearance63;
			this.txtPeso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Peso"));
			this.txtPeso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPeso.Location = new System.Drawing.Point(72, 40);
			this.txtPeso.Name = "txtPeso";
			this.txtPeso.Size = new System.Drawing.Size(100, 21);
			this.txtPeso.TabIndex = 1;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(16, 40);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(33, 16);
			this.label36.TabIndex = 58;
			this.label36.Text = "Peso:";
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(16, 16);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(46, 16);
			this.label41.TabIndex = 56;
			this.label41.Text = "Bulto / #";
			// 
			// txtBulto
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBulto.Appearance = appearance64;
			this.txtBulto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bulto"));
			this.txtBulto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBulto.Location = new System.Drawing.Point(72, 16);
			this.txtBulto.Name = "txtBulto";
			this.txtBulto.Size = new System.Drawing.Size(100, 21);
			this.txtBulto.TabIndex = 0;
			// 
			// cdvComprobante
			// 
			this.cdvComprobante.BindingContextCtrl = this;
			this.cdvComprobante.DataSet = this.cdsCompraTabla;
			this.cdvComprobante.TableName = "";
			this.cdvComprobante.TableViewName = "CompraComprobante";
			// 
			// cdvNumero
			// 
			this.cdvNumero.BindingContextCtrl = this;
			this.cdvNumero.DataSet = this.cdsCompraTabla;
			this.cdvNumero.TableName = "";
			this.cdvNumero.TableViewName = "CompraNumero";
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(602, 248);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 14;
			this.lblIva0.Text = "IVA 0";
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(696, 248);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 15;
			this.lblIvaSubt.Text = "IVA";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(632, 312);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 17;
			this.lblIva.Text = "IVA";
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(504, 288);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(58, 16);
			this.lblDescuento.TabIndex = 16;
			this.lblDescuento.Text = "Descuento";
			// 
			// imageList
			// 
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btTotal.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btTotal.Location = new System.Drawing.Point(504, 368);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(40, 23);
			this.btTotal.TabIndex = 20;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(240, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 30;
			this.label6.Text = "Fecha";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(408, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(27, 16);
			this.label7.TabIndex = 31;
			this.label7.Text = "Días";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(512, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 32;
			this.label8.Text = "Vence:";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(568, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 33;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(240, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(43, 16);
			this.lblCliente.TabIndex = 36;
			this.lblCliente.Text = "Cliente:";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(504, 264);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(50, 16);
			this.lblSubTotal.TabIndex = 37;
			this.lblSubTotal.Text = "SubTotal";
			// 
			// lblIce
			// 
			this.lblIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(504, 312);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(23, 16);
			this.lblIce.TabIndex = 38;
			this.lblIce.Text = "ICE";
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(624, 400);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(112, 23);
			this.lblAnulado.TabIndex = 41;
			this.lblAnulado.Text = "ANULADO";
			// 
			// CmbDescCod
			// 
			this.CmbDescCod.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 387;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Width = 179;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.CmbDescCod.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.CmbDescCod.DisplayMember = "Articulo";
			this.CmbDescCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.CmbDescCod.Location = new System.Drawing.Point(280, 112);
			this.CmbDescCod.Name = "CmbDescCod";
			this.CmbDescCod.Size = new System.Drawing.Size(152, 64);
			this.CmbDescCod.TabIndex = 53;
			this.CmbDescCod.ValueMember = "idArticulo";
			this.CmbDescCod.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Compra.Compra - DetCompra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance65.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance65;
			ultraGridColumn32.Header.VisiblePosition = 26;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.Caption = "Artículo";
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn34.Width = 216;
			ultraGridColumn35.Header.Caption = "Unidad";
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn35.Width = 53;
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 29;
			ultraGridColumn37.Hidden = true;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance66;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn38.NullText = "0";
			ultraGridColumn38.PromptChar = ' ';
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance67;
			ultraGridColumn39.Format = "";
			ultraGridColumn39.Header.VisiblePosition = 9;
			ultraGridColumn39.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn39.PromptChar = ' ';
			ultraGridColumn39.Width = 89;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance68;
			ultraGridColumn40.Header.Caption = "Iva";
			ultraGridColumn40.Header.VisiblePosition = 10;
			ultraGridColumn40.PromptChar = ' ';
			ultraGridColumn40.Width = 34;
			ultraGridColumn41.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn41.Header.VisiblePosition = 22;
			ultraGridColumn42.Header.Caption = "Ref. Codigo";
			ultraGridColumn42.Header.VisiblePosition = 20;
			ultraGridColumn43.Header.Caption = "Ref. Numero";
			ultraGridColumn43.Header.VisiblePosition = 21;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance69;
			ultraGridColumn44.Header.Caption = "Descuento";
			ultraGridColumn44.Header.VisiblePosition = 12;
			ultraGridColumn44.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn44.PromptChar = ' ';
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance70;
			ultraGridColumn45.Header.Caption = "Peso";
			ultraGridColumn45.Header.VisiblePosition = 19;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance71;
			ultraGridColumn46.Header.VisiblePosition = 28;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 18;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance72;
			ultraGridColumn48.Header.Caption = "% Desc.";
			ultraGridColumn48.Header.VisiblePosition = 13;
			ultraGridColumn48.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn48.PromptChar = ' ';
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance73;
			ultraGridColumn49.Header.VisiblePosition = 25;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 27;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 38;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 30;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 31;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 15;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 24;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.Caption = "Centro Costo";
			ultraGridColumn56.Header.VisiblePosition = 33;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn57.Header.Caption = "Proyecto";
			ultraGridColumn57.Header.VisiblePosition = 17;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn58.Header.VisiblePosition = 6;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 118;
			ultraGridColumn60.Header.VisiblePosition = 34;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance74;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 11;
			ultraGridColumn62.Header.Caption = "Importacion";
			ultraGridColumn62.Header.VisiblePosition = 35;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn63.Header.VisiblePosition = 36;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 5;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.Caption = "CentroCosto";
			ultraGridColumn65.Header.VisiblePosition = 14;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.Caption = "Proyecto";
			ultraGridColumn66.Header.VisiblePosition = 16;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 23;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 37;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 39;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 40;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 41;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 42;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 43;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 44;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 45;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 46;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 47;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 48;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.Caption = "Vendedor";
			ultraGridColumn79.Header.VisiblePosition = 49;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn80.Header.Caption = "Código";
			ultraGridColumn80.Header.VisiblePosition = 4;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn81.Header.Caption = "Unidades";
			ultraGridColumn81.Header.VisiblePosition = 8;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn81});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance75.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance76.FontData.BoldAsString = "True";
			appearance76.FontData.Name = "Arial";
			appearance76.FontData.SizeInPoints = 10F;
			appearance76.ForeColor = System.Drawing.Color.White;
			appearance76.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance76;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance77.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance77.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance77.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance77;
			appearance78.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance78.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance78;
			appearance79.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance79.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance79;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(720, 184);
			this.ultraGrid1.TabIndex = 8;
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
			// CmbCodDesc
			// 
			this.CmbCodDesc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Width = 170;
			ultraGridColumn84.Header.VisiblePosition = 2;
			ultraGridColumn84.Width = 332;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn82,
																											ultraGridColumn83,
																											ultraGridColumn84});
			this.CmbCodDesc.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.CmbCodDesc.DisplayMember = "Codigo";
			this.CmbCodDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.CmbCodDesc.Location = new System.Drawing.Point(80, 152);
			this.CmbCodDesc.Name = "CmbCodDesc";
			this.CmbCodDesc.Size = new System.Drawing.Size(176, 80);
			this.CmbCodDesc.TabIndex = 54;
			this.CmbCodDesc.ValueMember = "idArticulo";
			this.CmbCodDesc.Visible = false;
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
			// cmbCliente
			// 
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance80;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCliente"));
			ultraGridColumn85.Header.VisiblePosition = 0;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 1;
			ultraGridColumn86.Width = 325;
			ultraGridColumn87.Header.VisiblePosition = 2;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn85,
																											ultraGridColumn86,
																											ultraGridColumn87});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(288, 8);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(248, 21);
			this.cmbCliente.TabIndex = 2;
			this.cmbCliente.ValueMember = "idCliente";
			this.cmbCliente.Validated += new System.EventHandler(this.cmbCliente_Validated);
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
			// 
			// txtNumero
			// 
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance81;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(120, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 0;
			// 
			// dtFecha
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance82;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Fecha"));
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(288, 32);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 5;
			this.toolTip.SetToolTip(this.dtFecha, "Fecha de Transacción");
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// spnDias
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDias.Appearance = appearance83;
			this.spnDias.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnDias.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DiasCredito"));
			this.spnDias.Location = new System.Drawing.Point(440, 32);
			this.spnDias.MaskInput = "nnnn";
			this.spnDias.Name = "spnDias";
			this.spnDias.PromptChar = ' ';
			this.spnDias.Size = new System.Drawing.Size(64, 21);
			this.spnDias.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDias.TabIndex = 6;
			// 
			// dtVence
			// 
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtVence.Appearance = appearance84;
			this.dtVence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtVence.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaVencimiento"));
			dateButton4.Caption = "Today";
			this.dtVence.DateButtons.Add(dateButton4);
			this.dtVence.Format = "dd/MMM/yyyy";
			this.dtVence.Location = new System.Drawing.Point(560, 32);
			this.dtVence.Name = "dtVence";
			this.dtVence.NonAutoSizeHeight = 23;
			this.dtVence.Size = new System.Drawing.Size(112, 21);
			this.dtVence.SpinButtonsVisible = true;
			this.dtVence.TabIndex = 7;
			this.dtVence.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// spnBodega
			// 
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnBodega.Appearance = appearance85;
			this.spnBodega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bodega"));
			this.spnBodega.Location = new System.Drawing.Point(616, 8);
			this.spnBodega.MaskInput = "nnn";
			this.spnBodega.MinValue = 0;
			this.spnBodega.Name = "spnBodega";
			this.spnBodega.PromptChar = ' ';
			this.spnBodega.Size = new System.Drawing.Size(72, 21);
			this.spnBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnBodega.TabIndex = 3;
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(536, 8);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 23);
			this.btCliente.TabIndex = 76;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			this.btCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCliente_MouseDown);
			// 
			// cmbComprobante1
			// 
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComprobante1.Appearance = appearance86;
			this.cmbComprobante1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idComprobante"));
			this.cmbComprobante1.DataSource = this.cdvComprobante;
			ultraGridBand13.AddButtonCaption = "Cliente";
			ultraGridColumn88.Header.VisiblePosition = 2;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Width = 62;
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Width = 299;
			ultraGridColumn91.Header.VisiblePosition = 3;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 4;
			ultraGridColumn92.Hidden = true;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90,
																											ultraGridColumn91,
																											ultraGridColumn92});
			this.cmbComprobante1.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbComprobante1.DisplayMember = "Comprobante";
			this.cmbComprobante1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobante1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComprobante1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobante1.Location = new System.Drawing.Point(8, 32);
			this.cmbComprobante1.Name = "cmbComprobante1";
			this.cmbComprobante1.Size = new System.Drawing.Size(224, 21);
			this.cmbComprobante1.TabIndex = 4;
			this.cmbComprobante1.ValueMember = "idComprobante";
			this.cmbComprobante1.Validated += new System.EventHandler(this.cmbComprobante1_Validated);
			this.cmbComprobante1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbComprobante1_MouseUp);
			// 
			// cmbTipoCompra
			// 
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoCompra.Appearance = appearance87;
			this.cmbTipoCompra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoCompra.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idTipoFactura"));
			this.cmbTipoCompra.DataSource = this.cdvNumero;
			ultraGridBand14.AddButtonCaption = "Cliente";
			ultraGridColumn93.Header.VisiblePosition = 7;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn95.Header.VisiblePosition = 8;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 3;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 2;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 5;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 6;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 4;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 9;
			ultraGridColumn103.Header.VisiblePosition = 10;
			ultraGridColumn104.Header.VisiblePosition = 11;
			ultraGridColumn105.Header.VisiblePosition = 12;
			ultraGridColumn106.Header.VisiblePosition = 13;
			ultraGridColumn107.Header.VisiblePosition = 14;
			ultraGridColumn108.Header.VisiblePosition = 15;
			ultraGridColumn109.Header.VisiblePosition = 16;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn93,
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96,
																											ultraGridColumn97,
																											ultraGridColumn98,
																											ultraGridColumn99,
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108,
																											ultraGridColumn109});
			this.cmbTipoCompra.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbTipoCompra.DisplayMember = "Nombre";
			this.cmbTipoCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCompra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCompra.Enabled = false;
			this.cmbTipoCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCompra.Location = new System.Drawing.Point(8, 8);
			this.cmbTipoCompra.Name = "cmbTipoCompra";
			this.cmbTipoCompra.ReadOnly = true;
			this.cmbTipoCompra.Size = new System.Drawing.Size(112, 21);
			this.cmbTipoCompra.TabIndex = 78;
			this.cmbTipoCompra.ValueMember = "idTipoFactura";
			// 
			// btSaldo1
			// 
			this.btSaldo1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance88.Image = 5;
			this.btSaldo1.Appearance = appearance88;
			this.btSaldo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSaldo1.ImageList = this.imageList;
			this.btSaldo1.Location = new System.Drawing.Point(328, 400);
			this.btSaldo1.Name = "btSaldo1";
			this.btSaldo1.Size = new System.Drawing.Size(24, 23);
			this.btSaldo1.TabIndex = 11;
			this.toolTip.SetToolTip(this.btSaldo1, "Saldo de Artículo / Kardex de Articulo");
			this.btSaldo1.Click += new System.EventHandler(this.btSaldo_Click);
			this.btSaldo1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldo1_MouseUp);
			// 
			// btBuscaArticulo
			// 
			this.btBuscaArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance89.Image = 6;
			this.btBuscaArticulo.Appearance = appearance89;
			this.btBuscaArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscaArticulo.ImageList = this.imageList;
			this.btBuscaArticulo.Location = new System.Drawing.Point(352, 400);
			this.btBuscaArticulo.Name = "btBuscaArticulo";
			this.btBuscaArticulo.Size = new System.Drawing.Size(24, 23);
			this.btBuscaArticulo.TabIndex = 12;
			this.toolTip.SetToolTip(this.btBuscaArticulo, "Búsqueda de artículo seleccionado / Ultimos Movimientos");
			this.btBuscaArticulo.Click += new System.EventHandler(this.btBuscaArticulo_Click);
			this.btBuscaArticulo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btBuscaArticulo_MouseUp);
			// 
			// btTransforma
			// 
			this.btTransforma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance90.Image = 3;
			this.btTransforma.Appearance = appearance90;
			this.btTransforma.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTransforma.ImageList = this.imageList;
			this.btTransforma.Location = new System.Drawing.Point(376, 400);
			this.btTransforma.Name = "btTransforma";
			this.btTransforma.Size = new System.Drawing.Size(24, 23);
			this.btTransforma.TabIndex = 13;
			this.toolTip.SetToolTip(this.btTransforma, "Transformación / Distribuye Porcentajes");
			this.btTransforma.Click += new System.EventHandler(this.btTransforma_Click);
			this.btTransforma.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTransforma_MouseUp);
			// 
			// btNotaEntrega
			// 
			this.btNotaEntrega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance91.Image = 7;
			this.btNotaEntrega.Appearance = appearance91;
			this.btNotaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNotaEntrega.ImageList = this.imageList;
			this.btNotaEntrega.Location = new System.Drawing.Point(400, 400);
			this.btNotaEntrega.Name = "btNotaEntrega";
			this.btNotaEntrega.Size = new System.Drawing.Size(24, 23);
			this.btNotaEntrega.TabIndex = 14;
			this.toolTip.SetToolTip(this.btNotaEntrega, "Impresión de Nota de Entrega / Envío de Correo Electronico");
			this.btNotaEntrega.Click += new System.EventHandler(this.btNotaEntrega_Click);
			this.btNotaEntrega.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btNotaEntrega_MouseDown);
			// 
			// btRetencion
			// 
			this.btRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance92.Image = 0;
			this.btRetencion.Appearance = appearance92;
			this.btRetencion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRetencion.ImageList = this.imageList;
			this.btRetencion.Location = new System.Drawing.Point(424, 400);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(24, 23);
			this.btRetencion.TabIndex = 15;
			this.toolTip.SetToolTip(this.btRetencion, "Retención en la fuente");
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			this.btRetencion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRetencion_MouseDown);
			// 
			// btCreaAsiento
			// 
			this.btCreaAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance93.Image = 9;
			this.btCreaAsiento.Appearance = appearance93;
			this.btCreaAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCreaAsiento.Location = new System.Drawing.Point(448, 400);
			this.btCreaAsiento.Name = "btCreaAsiento";
			this.btCreaAsiento.Size = new System.Drawing.Size(24, 23);
			this.btCreaAsiento.TabIndex = 16;
			this.btCreaAsiento.Text = "A";
			this.toolTip.SetToolTip(this.btCreaAsiento, "Crea Asiento");
			this.btCreaAsiento.Click += new System.EventHandler(this.btCreaAsiento_Click);
			this.btCreaAsiento.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCreaAsiento_MouseDown);
			// 
			// btGuiaRemision
			// 
			this.btGuiaRemision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance94.Image = 8;
			this.btGuiaRemision.Appearance = appearance94;
			this.btGuiaRemision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGuiaRemision.ImageList = this.imageList;
			this.btGuiaRemision.Location = new System.Drawing.Point(472, 400);
			this.btGuiaRemision.Name = "btGuiaRemision";
			this.btGuiaRemision.Size = new System.Drawing.Size(24, 23);
			this.btGuiaRemision.TabIndex = 17;
			this.toolTip.SetToolTip(this.btGuiaRemision, "Guía de Remisión");
			this.btGuiaRemision.Visible = false;
			this.btGuiaRemision.Click += new System.EventHandler(this.btGuiaRemision_Click);
			this.btGuiaRemision.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGuiaRemision_MouseDown);
			// 
			// btActualizaPrecio
			// 
			this.btActualizaPrecio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance95.Image = 11;
			this.btActualizaPrecio.Appearance = appearance95;
			this.btActualizaPrecio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btActualizaPrecio.ImageList = this.imageList;
			this.btActualizaPrecio.Location = new System.Drawing.Point(496, 400);
			this.btActualizaPrecio.Name = "btActualizaPrecio";
			this.btActualizaPrecio.Size = new System.Drawing.Size(24, 23);
			this.btActualizaPrecio.TabIndex = 18;
			this.toolTip.SetToolTip(this.btActualizaPrecio, "Actualiza Precios");
			this.btActualizaPrecio.Visible = false;
			this.btActualizaPrecio.Click += new System.EventHandler(this.btActualizaPrecio_Click);
			this.btActualizaPrecio.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btActualizaPrecio_MouseDown);
			// 
			// btNuevo
			// 
			this.btNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNuevo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNuevo.ImageList = this.imageList;
			this.btNuevo.Location = new System.Drawing.Point(520, 400);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(24, 23);
			this.btNuevo.TabIndex = 19;
			this.btNuevo.Text = "&N";
			this.toolTip.SetToolTip(this.btNuevo, "Nueva Línea");
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idAsiento"));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(512, 384);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(48, 21);
			this.txtIdAsiento.TabIndex = 88;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Location = new System.Drawing.Point(600, 400);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 21);
			this.txtIdCompra.TabIndex = 89;
			// 
			// txtBorrar
			// 
			this.txtBorrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBorrar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Borrar"));
			this.txtBorrar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBorrar.Location = new System.Drawing.Point(624, 376);
			this.txtBorrar.Name = "txtBorrar";
			this.txtBorrar.Size = new System.Drawing.Size(48, 21);
			this.txtBorrar.TabIndex = 90;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Compra";
			this.barraDato1.DataNombreId = "idCompra";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCompra;
			this.barraDato1.DataTabla = "Compra";
			this.barraDato1.DataTablaHija = "DetCompra";
			this.barraDato1.Location = new System.Drawing.Point(8, 400);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 10;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance96;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl2);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl3);
			this.utcTransaccion.Controls.Add(this.pgEntrega);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl5);
			this.utcTransaccion.Controls.Add(this.TabPageCOA);
			this.utcTransaccion.Controls.Add(this.tbpAutorizacion);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl6);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl4);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 256);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(488, 136);
			this.utcTransaccion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcTransaccion.TabIndex = 9;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Notas";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Pago";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Desc.";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Entrega";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Cuotas";
			ultraTab6.TabPage = this.TabPageCOA;
			ultraTab6.Text = "A.T.";
			ultraTab7.TabPage = this.tbpAutorizacion;
			ultraTab7.Text = "Autoriz.";
			ultraTab8.TabPage = this.ultraTabPageControl6;
			ultraTab8.Text = "Busca";
			ultraTab9.TabPage = this.ultraTabPageControl4;
			ultraTab9.Text = "Otro";
			this.utcTransaccion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3,
																																														 ultraTab4,
																																														 ultraTab5,
																																														 ultraTab6,
																																														 ultraTab7,
																																														 ultraTab8,
																																														 ultraTab9});
			this.utcTransaccion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(486, 115);
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance97;
			this.txtDescIva0Total.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento0"));
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(560, 288);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.ReadOnly = true;
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 21);
			this.txtDescIva0Total.TabIndex = 112;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance98;
			this.txtDescIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento"));
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(640, 288);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.ReadOnly = true;
			this.txtDescIvaTotal.Size = new System.Drawing.Size(88, 21);
			this.txtDescIvaTotal.TabIndex = 111;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance99;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalIva"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(640, 264);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.ReadOnly = true;
			this.txtIva.Size = new System.Drawing.Size(88, 21);
			this.txtIva.TabIndex = 110;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance100;
			this.txtIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalExcento"));
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(560, 264);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.ReadOnly = true;
			this.txtIva0.Size = new System.Drawing.Size(80, 21);
			this.txtIva0.TabIndex = 109;
			// 
			// txtIceTotal
			// 
			this.txtIceTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIceTotal.Appearance = appearance101;
			this.txtIceTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Ice"));
			this.txtIceTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIceTotal.FormatString = "#,##0.00";
			this.txtIceTotal.Location = new System.Drawing.Point(552, 312);
			this.txtIceTotal.Name = "txtIceTotal";
			this.txtIceTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIceTotal.PromptChar = ' ';
			this.txtIceTotal.Size = new System.Drawing.Size(64, 21);
			this.txtIceTotal.TabIndex = 115;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance102;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Total"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(552, 360);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.ReadOnly = true;
			this.txtTotal.Size = new System.Drawing.Size(176, 34);
			this.txtTotal.TabIndex = 114;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance103;
			this.txtIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Iva"));
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(656, 312);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.ReadOnly = true;
			this.txtIvaTotal.Size = new System.Drawing.Size(72, 21);
			this.txtIvaTotal.TabIndex = 113;
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
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(440, 152);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(200, 56);
			this.cmbProyecto1.TabIndex = 116;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			// 
			// cmbSubProyecto1
			// 
			this.cmbSubProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto1.DisplayMember = "Nombre";
			this.cmbSubProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto1.Location = new System.Drawing.Point(504, 88);
			this.cmbSubProyecto1.Name = "cmbSubProyecto1";
			this.cmbSubProyecto1.Size = new System.Drawing.Size(232, 48);
			this.cmbSubProyecto1.TabIndex = 117;
			this.cmbSubProyecto1.ValueMember = "idSubProyecto";
			this.cmbSubProyecto1.Visible = false;
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cmbUnidad
			// 
			this.cmbUnidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUnidad.DataMember = "ArticuloUnidad";
			this.cmbUnidad.DataSource = this.cdsUnidad;
			ultraGridBand15.AddButtonCaption = "CompraComprobante";
			ultraGridColumn110.Header.VisiblePosition = 3;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 2;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 1;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 4;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 5;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 0;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113,
																											ultraGridColumn114,
																											ultraGridColumn115});
			this.cmbUnidad.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbUnidad.DisplayMember = "Unidades";
			this.cmbUnidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUnidad.Location = new System.Drawing.Point(576, 168);
			this.cmbUnidad.Name = "cmbUnidad";
			this.cmbUnidad.Size = new System.Drawing.Size(136, 64);
			this.cmbUnidad.TabIndex = 118;
			this.cmbUnidad.ValueMember = "idUnidad";
			this.cmbUnidad.Visible = false;
			// 
			// cdsUnidad
			// 
			this.cdsUnidad.BindingContextCtrl = this;
			this.cdsUnidad.DataLibrary = "LibFacturacion";
			this.cdsUnidad.DataLibraryUrl = "";
			this.cdsUnidad.DataSetDef = "dsArticuloTabla";
			this.cdsUnidad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUnidad.Name = "cdsUnidad";
			this.cdsUnidad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsUnidad.SchemaDef = null;
			this.cdsUnidad.BeforeFill += new C1.Data.FillEventHandler(this.cdsUnidad_BeforeFill);
			// 
			// txtNombre
			// 
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance104;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(288, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(232, 21);
			this.txtNombre.TabIndex = 1;
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			// 
			// btBarras
			// 
			this.btBarras.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance105.Image = ((object)(resources.GetObject("appearance105.Image")));
			this.btBarras.Appearance = appearance105;
			this.btBarras.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBarras.Location = new System.Drawing.Point(544, 400);
			this.btBarras.Name = "btBarras";
			this.btBarras.Size = new System.Drawing.Size(24, 23);
			this.btBarras.TabIndex = 119;
			this.btBarras.Click += new System.EventHandler(this.btBarras_Click);
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance106.Image = ((object)(resources.GetObject("appearance106.Image")));
			this.btAuditoria.Appearance = appearance106;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(568, 400);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 24);
			this.btAuditoria.TabIndex = 121;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// txtServicio
			// 
			this.txtServicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance107.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtServicio.Appearance = appearance107;
			this.txtServicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Servicio"));
			this.txtServicio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtServicio.FormatString = "#,##0.00";
			this.txtServicio.Location = new System.Drawing.Point(552, 336);
			this.txtServicio.Name = "txtServicio";
			this.txtServicio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtServicio.PromptChar = ' ';
			this.txtServicio.Size = new System.Drawing.Size(64, 21);
			this.txtServicio.TabIndex = 123;
			// 
			// lblServicio
			// 
			this.lblServicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblServicio.AutoSize = true;
			this.lblServicio.Location = new System.Drawing.Point(504, 336);
			this.lblServicio.Name = "lblServicio";
			this.lblServicio.Size = new System.Drawing.Size(45, 16);
			this.lblServicio.TabIndex = 122;
			this.lblServicio.Text = "Servicio";
			// 
			// lblPropina
			// 
			this.lblPropina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPropina.AutoSize = true;
			this.lblPropina.Location = new System.Drawing.Point(616, 336);
			this.lblPropina.Name = "lblPropina";
			this.lblPropina.Size = new System.Drawing.Size(43, 16);
			this.lblPropina.TabIndex = 124;
			this.lblPropina.Text = "Propina";
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bodega"));
			ultraGridColumn116.Header.VisiblePosition = 1;
			ultraGridColumn117.Header.VisiblePosition = 0;
			ultraGridColumn117.Width = 180;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn116,
																											ultraGridColumn117});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(576, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 126;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// lblTiempo
			// 
			this.lblTiempo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblTiempo.AutoSize = true;
			this.lblTiempo.Location = new System.Drawing.Point(504, 250);
			this.lblTiempo.Name = "lblTiempo";
			this.lblTiempo.Size = new System.Drawing.Size(17, 16);
			this.lblTiempo.TabIndex = 127;
			this.lblTiempo.Text = "T: ";
			// 
			// cmbBodegaLinea
			// 
			this.cmbBodegaLinea.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaLinea.DisplayMember = "Nombre";
			this.cmbBodegaLinea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaLinea.Location = new System.Drawing.Point(56, 88);
			this.cmbBodegaLinea.Name = "cmbBodegaLinea";
			this.cmbBodegaLinea.Size = new System.Drawing.Size(176, 72);
			this.cmbBodegaLinea.TabIndex = 128;
			this.cmbBodegaLinea.ValueMember = "Bodega";
			this.cmbBodegaLinea.Visible = false;
			// 
			// txtPropina
			// 
			this.txtPropina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance108.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPropina.Appearance = appearance108;
			this.txtPropina.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Propina"));
			this.txtPropina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPropina.FormatString = "#,##0.00";
			this.txtPropina.Location = new System.Drawing.Point(656, 336);
			this.txtPropina.Name = "txtPropina";
			this.txtPropina.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPropina.PromptChar = ' ';
			this.txtPropina.Size = new System.Drawing.Size(72, 21);
			this.txtPropina.TabIndex = 129;
			// 
			// cmbUnidadB
			// 
			this.cmbUnidadB.Cursor = System.Windows.Forms.Cursors.SizeWE;
			this.cmbUnidadB.DataMember = "ArticuloUnidad";
			this.cmbUnidadB.DataSource = this.cdsUnidad;
			ultraGridBand17.AddButtonCaption = "CompraComprobante";
			ultraGridColumn118.Header.VisiblePosition = 3;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 2;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 1;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 4;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 5;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 0;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn118,
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121,
																											ultraGridColumn122,
																											ultraGridColumn123});
			this.cmbUnidadB.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbUnidadB.DisplayMember = "Unidades";
			this.cmbUnidadB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUnidadB.Location = new System.Drawing.Point(288, 178);
			this.cmbUnidadB.Name = "cmbUnidadB";
			this.cmbUnidadB.Size = new System.Drawing.Size(136, 72);
			this.cmbUnidadB.TabIndex = 130;
			this.cmbUnidadB.ValueMember = "idUnidad";
			this.cmbUnidadB.Visible = false;
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
			// cmbVendedorGrilla
			// 
			this.cmbVendedorGrilla.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn124.Header.VisiblePosition = 0;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 1;
			ultraGridColumn125.Width = 387;
			ultraGridColumn126.Header.VisiblePosition = 2;
			ultraGridColumn126.Width = 179;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn124,
																											ultraGridColumn125,
																											ultraGridColumn126});
			this.cmbVendedorGrilla.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbVendedorGrilla.DisplayMember = "";
			this.cmbVendedorGrilla.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedorGrilla.Location = new System.Drawing.Point(440, 136);
			this.cmbVendedorGrilla.Name = "cmbVendedorGrilla";
			this.cmbVendedorGrilla.Size = new System.Drawing.Size(152, 64);
			this.cmbVendedorGrilla.TabIndex = 131;
			this.cmbVendedorGrilla.ValueMember = "";
			this.cmbVendedorGrilla.Visible = false;
			// 
			// btDigitalizacion
			// 
			this.btDigitalizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDigitalizacion.ImageIndex = 0;
			this.btDigitalizacion.ImageList = this.imageList;
			this.btDigitalizacion.Location = new System.Drawing.Point(592, 400);
			this.btDigitalizacion.Name = "btDigitalizacion";
			this.btDigitalizacion.Size = new System.Drawing.Size(24, 23);
			this.btDigitalizacion.TabIndex = 132;
			this.btDigitalizacion.ToolTip = "Digitalización";
			this.btDigitalizacion.Visible = false;
			this.btDigitalizacion.Click += new System.EventHandler(this.btDigitalizacion_Click);
			this.btDigitalizacion.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btDigitalizacion_MouseDown);
			// 
			// Compra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(744, 429);
			this.Controls.Add(this.btDigitalizacion);
			this.Controls.Add(this.cmbVendedorGrilla);
			this.Controls.Add(this.cmbUnidadB);
			this.Controls.Add(this.txtPropina);
			this.Controls.Add(this.cmbBodegaLinea);
			this.Controls.Add(this.lblTiempo);
			this.Controls.Add(this.lblPropina);
			this.Controls.Add(this.lblServicio);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtServicio);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.btBarras);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbUnidad);
			this.Controls.Add(this.cmbSubProyecto1);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.txtIceTotal);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.utcTransaccion);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtBorrar);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btActualizaPrecio);
			this.Controls.Add(this.btGuiaRemision);
			this.Controls.Add(this.btCreaAsiento);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.btNotaEntrega);
			this.Controls.Add(this.btTransforma);
			this.Controls.Add(this.btBuscaArticulo);
			this.Controls.Add(this.btSaldo1);
			this.Controls.Add(this.cmbTipoCompra);
			this.Controls.Add(this.cmbComprobante1);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.dtVence);
			this.Controls.Add(this.spnDias);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.CmbCodDesc);
			this.Controls.Add(this.CmbDescCod);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.btTotal);
			this.helpProvider1.SetHelpKeyword(this, "Facturación.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.Name = "Compra";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Compra";
			this.Resize += new System.EventHandler(this.Compra_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Compra_Closing);
			this.Load += new System.EventHandler(this.Compra2_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnVuelto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadoCred)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFlete)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibeFabara)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprob1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRevision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteresPagar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpezarEn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagosCada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			this.TabPageCOA.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbIce)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.OptBienServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefrendo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).EndInit();
			this.tbpAutorizacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPrioridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrBusca)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.spnNumFrecCopiarCotiz)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumCopiarCotiz)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvConfiteca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfiteca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBulto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbDescCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbCodDesc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).EndInit();
			this.utcTransaccion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIceTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPropina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadB)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedorGrilla)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales
		private Acceso miAcceso;
		int IdTipoFactura = 1;
		int IdBusqueda = 0;
		private C1.Data.C1DataRow drCompra;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		bool bAutoImpresor = false;
		bool bNumeraImprimir = false;
		private bool bLote = false;
		private int IdSigno = 2;

		int IProyecto = 0;
		int ISubProyecto = 0;
		int IBodega = 0;
		struct Proveedor
		{
			public DateTime FechaCaducidad;
			public string SerieFactura;
			public string AutorizaImprenta;
			public string AutorizaFactura;
			public int DiasCredito;
			public int idVendedor;
			public int idCredTribut;
			public double MaximoCredito;
			public double DescuentoCredito;
			public double DescuentoContado;
			public int idTipoPrecio;
			public bool InstPublica;

			public Proveedor(C1.Data.C1DataSet ds, int idCliente)
			{
				string stSelect = "Select IsNull(FechaCaducidad, '19000101'), IsNull(SerieFactura, ''), "
					+ "IsNull(AutorizaImprenta, ''), IsNull(AutorizaFactura, ''), "
					+ "IsNull(DiasCredito, 0), IsNull(idVendedor, 0), "
					+ "IsNull(idCredTribut, 0), IsNull(MaximoCredito, 0), "
					+ "IsNull(DescuentoCredito, 0), IsNull(DescuentoContado, 0), "
					+ "IsNull(idTipoPrecio, 0), IsNull(instPublica, 0) "
					+ "From Cliente Where idCliente = " + idCliente.ToString();
				SqlDataReader miReader = null;
				this.FechaCaducidad = DateTime.Today;
				this.SerieFactura = "001001";
				this.AutorizaImprenta = "";
				this.AutorizaFactura = "";
				this.DiasCredito = 0;
				this.idVendedor = 0;
				this.idCredTribut = 1;
				this.MaximoCredito = 0;
				this.DescuentoCredito = 0;
				this.DescuentoContado = 0;
				this.idTipoPrecio = 1;
				this.InstPublica = false;
				try
				{
					miReader = Funcion.rEscalarSQL(ds, stSelect, true);
					miReader.Read();
					this.FechaCaducidad = miReader.GetDateTime(0);
					this.SerieFactura = miReader.GetString(1);
					this.AutorizaImprenta = miReader.GetString(2);
					this.AutorizaFactura = miReader.GetString(3);
					this.DiasCredito = miReader.GetInt32(4);
					this.idVendedor = miReader.GetInt32(5);
					this.idCredTribut = miReader.GetInt32(6);
					this.MaximoCredito = miReader.GetDouble(7);
					this.DescuentoCredito = miReader.GetDouble(8);
					this.DescuentoContado = miReader.GetDouble(9);
					this.idTipoPrecio = miReader.GetInt32(10);
					this.InstPublica = miReader.GetBoolean(11);
				}
				catch{}
				finally
				{
					miReader.Close();
				}
			}
		}

		private Proveedor miProveedor;
		private bool bIncluyeIva = false;
		public DateTime FechaCaducidadSri;
		bool bListaCliente = true;
		string controlW="";
		string NuevoCodigo="";
		string controlW1="";		
		private bool bActualiza = true;
		private bool bServicioAuto = false;
		CultureInfo us = new CultureInfo("en-US");
		private int iCuentaTotal = 0;
		int ProvRecibFabara = 0;
		bool bRastreo = false;
		bool bVendedor = false;
		#endregion
		#region Ventana Compra
		public Compra(int idTransaccion)
		{
			InitializeComponent();
			if (idTransaccion > 0)
				IdTipoFactura = idTransaccion;
			else
			{
				IdBusqueda = -idTransaccion;
				string stTransaccion = "Select c.idTipoFactura From Compra c where c.idCompra = " + IdBusqueda.ToString();
				IdTipoFactura = Funcion.iEscalarSQL(cdsCompra, stTransaccion);
			}
		}

		public Compra(int idTransaccion, int idBusca)
		{
			InitializeComponent();
			IdTipoFactura = idTransaccion;
			IdBusqueda = idBusca;
		}

		private void Compra2_Load(object sender, System.EventArgs e)
		{
			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");
			Cursor = Cursors.Default;
			Seguridades();
			Configuracion();
			Personalizaciones();
			Visibilidad();
			#region Auditoria
			int iLong = cmbTipoCompra.Text.Trim().Length;
			if (iLong > 20) iLong = 20;
			string stAudita = "Exec AuditaCrear 53, 6, '" + cmbTipoCompra.Text.Trim().Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			#endregion Auditoria
			#region Busqueda de Registro
			if (IdBusqueda > 0)
			{
				barraDato1.IdRegistro = IdBusqueda;
				barraDato1.ProximoId(5);
			}
			else
				barraDato1.PosUltima();
			#endregion Busqueda de Registro
			#region Habilita controles
			barraDato1.HabilitaControles(false);
			HabilitaTab();
			ModificaGrilla(false);
			#region habilita campos para copias de cotizacion 
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRCCTZ'")) 				
			{		
				if(IdTipoFactura ==3)
				{
					spnNumCopiarCotiz.Visible = true;
					btnCopiarCotiz.Visible = true;
					spnNumFrecCopiarCotiz.Visible = true;
					spnNumCopiarCotiz.Enabled = true;
					spnNumFrecCopiarCotiz.Enabled = true;
				}
			}
			#endregion habilita campos para copias dee cotizacion
			#endregion habilita controles
			#region Titulo y texto proveed. en Orden Compra, Anticipo Proveedor, Reserva Proveedor, Compra, Devolución Compra, Ingreso Bodega, Pedido Proveedor
			if (IdTipoFactura == 2 || IdTipoFactura == 4 || IdTipoFactura == 25
				|| IdTipoFactura == 26 || IdTipoFactura == 6 || IdTipoFactura == 9
				|| IdTipoFactura == 36 || IdTipoFactura == 28 || IdTipoFactura == 14)
			{
				lblCliente.Text = "Proveed.";
				ProvRecibFabara = 1;
			}
			Text = cmbTipoCompra.Text;
			#endregion 

			LlenaCombos();
			#region Autoimpresor
			//			stSelect = "SELECT Count(*) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
			//				+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString();
			//			int iCuentaAutoImp = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
			//			if (iCuentaAutoImp == 1)
			//			{
			//				stSelect = "SELECT IsNull(AutoImpresor, 0) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
			//					+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString();
			//				bAutoImpresor = Funcion.bEjecutaSQL(cdsSeteoF, stSelect);
			//			}
			//			// En ventas hay nota de venta y factura
			//			if (iCuentaAutoImp == 2 && IdTipoFactura == 1)
			//			{
			//				stSelect = "SELECT IsNull(AutoImpresor, 0) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
			//					+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString() + " And idDocumento In (1, 2)";
			//				bAutoImpresor = Funcion.bEjecutaSQL(cdsSeteoF, stSelect);
			//				iCuentaAutoImp = 1;
			//			}
			//			if (iCuentaAutoImp > 1)
			//			{
			//				MessageBox.Show("Numeracion de autoimpresor esta duplicada", "Información");
			//			}
			//
			//			if (bAutoImpresor)
			//			{
			//				dtFecha.ReadOnly = true;
			//				txtNumero.ReadOnly = true;
			//				cmbCaducidad.ReadOnly = true;
			//				txtAutFactura.ReadOnly = true;
			//				string stExecCaducado = "Exec SriAutorizacionCaducada " + IdTipoFactura.ToString();
			//
			//				int iDiasCaducidad = Funcion.iEscalarSQL(cdsSeteoF, stExecCaducado, true);
			//				FechaCaducidadSri = DateTime.Today.AddDays(iDiasCaducidad);
			//				if (iDiasCaducidad <= 0)
			//				{
			//					MessageBox.Show("Autorizacion Caducada", "Información");
			//					barraDato1.VisibleNuevo = false;
			//					Close();
			//				}
			//				if (iDiasCaducidad < 30)
			//				{
			//					MessageBox.Show("Faltan " + iDiasCaducidad.ToString()
			//						+ " dias para que caduque autorizacion");
			//				}
			//			}
			#endregion Autoimpresor

			// Valores por omision de Grilla al crear nueva linea
			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'I'";
			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
			{
				ultraTextEditor4.ReadOnly = true;
				label14.Text= "BIO";
			}
		}
		
		private void LlenaCombos()
		{
			#region Vendedor
			string stExec = "Exec PersonalVendedor " + IdTipoFactura;
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbVendedor.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
			#endregion Vendedor
			#region Proyecto
			stExec = "Exec ProyectoCarga " + IdTipoFactura;
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbProyecto1.DataSource = cmbProyecto.DataSource;
			cmbProyecto.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;
			cmbProyecto1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;
			#endregion Proyecto
			#region SubProyecto
			stExec = "Exec SubProyectoCarga " + IdTipoFactura;
			cmbSubProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbSubProyecto1.DataSource = cmbSubProyecto.DataSource;
			cmbSubProyecto.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;
			cmbSubProyecto1.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;
			#endregion SubProyecto
			#region Lista de Precios
			stExec = "select idTipoPrecio, Precio from ClientePrecio";
			cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbPrecio.DisplayLayout.Bands[0].Columns["idTipoPrecio"].Hidden = true;
			#endregion Lista de Precios

			cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga");
			ArticuloCarga();
			ClienteCarga();
		}

		private void ArticuloCarga()
		{
			if (drSeteoF["VerDesplegable"] != DBNull.Value && (bool) drSeteoF["VerDesplegable"])
			{
				string stExec = "Exec ArticuloCargaNombre " + IdTipoFactura;
				CmbDescCod.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				CmbDescCod.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;

				stExec = "Exec ArticuloCargaCodigo " + IdTipoFactura;
				CmbCodDesc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				CmbCodDesc.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
			}
		}

		private void ClienteCarga()
		{
			if (drSeteoF["VerNombreLista"] != DBNull.Value && (bool) drSeteoF["VerNombreLista"])
			{
				string stExec = "Exec ClienteCargaNombre " + IdTipoFactura;
				cmbCliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				cmbCliente.DisplayLayout.Bands[0].Columns["idCliente"].Hidden = true;
				txtNombre.Width = 0;
			}
			else
				cmbCliente.Width = 0;

			cmbRecibe.DataSource = 	Funcion.dvProcedimiento(cdsCompra,
				"Select Nombre, Codigo, idCliente From Cliente Where Entrega = 1 and (Proveedor = "+ProvRecibFabara+" or Ambos = 1)");
		}
		private void Personalizaciones()
		{
			#region Binaria Vendedor
	
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FCVB'")) 				
			{
				if (IdTipoFactura == 4)return;
				bVendedor = true;
				string stSelect = "select idpersonal,Nombre from Personal where Vendedor =1";

				cmbVendedorGrilla.DataSource  = Funcion.dvProcedimiento(cdsCompra,stSelect);
				cmbVendedorGrilla.DisplayMember = "Nombre";
				cmbVendedorGrilla.ValueMember = "idPersonal";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idVendedor"].Hidden = false;										

			}
			#endregion Binaria Vendedor
			#region Cambios Botrosa
			///////////////////////////////////////////////////////////
			//Cambios para Botrosa 
			//agrandar el campo de bodegas para ver el nombre			
			// agrandar campo de vendedor
			//Filtrar en articulos que se escoja solo los de Inventarios
			//en Ingresos y egresos de BOdega
			//									if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPAI'")) 
			//									{if (IdTipoFactura == 8 || IdTipoFactura == 9)
			//										{ this.cdvArtCodigo.RowFilter = "idTipoGrupo In (1)";
			//											this.cdvArticulo.RowFilter = "idTipoGrupo In (1)";									}									}
			//poner como obligatorio ingresar Centro de costo: en la funcion grabar llamar a funcion CambiosBotrosa
			////funcoion para proyectos por centro de costo
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRPBP'")) 				
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProyectoBR"].Hidden = false;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CentroCostoBr"].Hidden = false;
			}
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVPSA'")) 				
				bServicioAuto = true;
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRPB'")) 				
			{				
				//crear en el grid UnidadBotrosa y un crear un cmbUnidadB y enlazar 
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["UnidadBotrosa"].Hidden = false;						
				//craer un cmbBodegaLinea se asigna y se hace el enlace 			
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Bodega"].ValueList = cmbBodegaLinea;
			}
			#endregion
			#region Confiteca
			//////////////////////////////////////////////////////////
			///Comisiones Externas para Confiteca
			
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FCOMEXT'")) 				
			{
				lblComisionExt.Visible = true;
				cmbNombre.Visible = true;
				btnComisionExt.Visible=true;
			}
			//////////////////////////////////////////////////////////			
			#endregion
			#region STI
			//////////////////////////////////////////////////////////
			/////cambio STI
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRSTI'"))
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["RefNumero"].Header.Caption = "Numero OP";
			}
			#endregion
			#region Academico
			if (MenuLatinium.bAcademica)
			{
				string stCuenta = "Select Count(*) From Compra";
				iCuentaTotal = Funcion.iEscalarSQL(cdsSeteoF, stCuenta, true);
			}
			#endregion Academico
			#region Fabara
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FRECIBEF'")) 				
			{
				cmbRecibeFabara.Visible=true;
				cmbRecibe.Visible=false;

				cmbRecibeFabara.DataSource = Funcion.dvProcedimiento(cdsCompra,
					"Select Nombre, Codigo, idCliente From Cliente Where Entrega = 1 and (Proveedor = "+ProvRecibFabara+" or Ambos =1)");
				cmbRecibeFabara.DisplayMember = "Nombre";
				cmbRecibeFabara.ValueMember = "idCliente";

			}
			#endregion Fabara
		}

		private void Visibilidad()
		{
			cmbBodega.Width = 0;
			txtIdAsiento.Width = 0;
			txtIdCompra.Width = 0;
			if (!miAcceso.Autorizacion) chkAutPresupuestaria.Visible = false;
			if (!miAcceso.Buscar) barraDato1.BarraMovimiento = false;
			if (miAcceso.ListaPrecios) 
			{
				cmbPrecio.Visible = true;
				btCambiarPrecio.Visible = true;
			}
			#region Bodegas por nombre
			if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"])
			{
				string stExecBodega = "Exec BodegaUsuario "	+ IdTipoFactura.ToString();
				cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExecBodega);

				if (cmbBodega.Rows.Count > 0)
				{
					cmbBodega.Width = 110;
					cmbBodega.Left = lblBodega.Left-6;
					spnBodega.Width = 0;
					lblBodega.Visible=false;
				}
				else
				{
					MessageBox.Show("No hay bodegas definidas en Herramientas - Seguridad - Nombre de Bodega", "Información");
				}
			}
			#endregion Lista de Bodegas
			VisibleBotones();
			if (drSeteoF["VerServicio"] == DBNull.Value || (bool) drSeteoF["VerServicio"] == false)
				MoverTotales();
		}

		private void Seguridades()
		{
			#region Codigo de Seguridad
			string stTipo = "04";
			switch( this.IdTipoFactura)
			{
				case 1:
				case 11:
				case 27:
					stTipo += "03";
					break;
				case 3:
					stTipo += "04";
					break;
				case 4:
				case 28:
					stTipo += "01";
					break;
				case 2:
				case 5:
				case 6:
					stTipo += "0" + this.IdTipoFactura.ToString().Trim();
					break;
				case 8:
				case 9:
					int idT = IdTipoFactura - 1;
					stTipo += "0" + idT.ToString().Trim();
					break;
				case 10:
					stTipo = "MP04";
					break;
				case 12:
					stTipo = "MP03";
					break;
				case 31:
					stTipo = "0408";
					break;
				case 36:
					stTipo = "040208";
					break;
				case 99:
					stTipo = "04090L9";
					break;
				default:
					stTipo = "0401";
					break;
			}
			miAcceso =new Acceso(cdsSeteoF, stTipo);
			#endregion Tipo de Factura
			#region Cambiar Fecha
			if (!miAcceso.CambiarFecha) 
			{
				dtFecha.ReadOnly = true;
				toolTip.SetToolTip(dtFecha, "Fecha desactivada en seguridades del usuario. Consulte al administrador del sistema");
			}
			#endregion Cambiar Fecha
			#region Cambiar Numeracion
			if (!miAcceso.CambiarNumeracion) 
			{
				toolTip.SetToolTip(txtNumero, "Numeración desactivada en seguridades del usuario. Consulte al administrador del sistema");
				txtNumero.ReadOnly = true;
			}
			#endregion Cambiar Numeracion
			#region Cambiar Precio
			// ! = Not && = And || = Or 
			if (!miAcceso.CambiarPrecio)// && (IdTipoFactura == 1 || IdTipoFactura == 3))
				ultraGrid1.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.Disabled;
			#endregion Cambio de Precio
			#region Cambiar Descuento
			if (!miAcceso.CambiarDescuento && (IdTipoFactura == 1 || IdTipoFactura == 3))
			{
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoArt"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				txtDescIva.ReadOnly = true;
				txtDescIva0.ReadOnly = true;
				txtDescPorcIva.ReadOnly = true;
				txtDescPorcIva0.ReadOnly = true;
			}
			#endregion Cambio de Descuento
		}
		private void Configuracion()
		{
			#region Asigna Signo
			string stSelect = "SELECT Signo FROM CompraNumero WHERE idTipoFactura = " + IdTipoFactura;
			IdSigno = Funcion.iEscalarSQL(cdsCompra, stSelect);
			#endregion Asigna Signo
			barraDato1.stFiltroExt = " And idTipoFactura = " + IdTipoFactura;
			#region Lectura de Archivo - Propiedades - Configuracion
			drSeteoC = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#endregion
			#region Incluye Iva en venta cotizacion dev de venta ingreso y egreso
			if ((bool) drSeteoF["IncluyeIva"] && (IdTipoFactura == 1 
				|| IdTipoFactura == 3 || IdTipoFactura == 5 
				|| IdTipoFactura == 8 || IdTipoFactura == 9)) bIncluyeIva = true;
			#endregion Incluye Iva
			CargaProyecto();
			#region Numeracion al Imprimir en ventas
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRNBG'")) 
				if (IdTipoFactura == 1)
					bNumeraImprimir = true;
			#endregion Numera al grabar
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

		private void VisibleBotones()
		{
			if (IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 9) // Compra, Egreso, Ingreso
			{
				if ((bool)drSeteoF["CompraUltimaC"] && !(bool) drSeteoF["ActualizarCosto"])
					btActualizaPrecio.Visible = true;
			}
			switch (IdTipoFactura)
			{
				case 1: // Venta 
				case 6: // Devolucion en compra
				case 4: // Compra
				case 5: // Devolucion en venta
				case 8: // Ingreso de Bodega
				case 9: // Egreso de Bodega
				case 10: // Consumo Real
				case 12: // Producto Terminado
				case 27: // Nota debito venta
				case 11: // Reservacion Cliente
				case 26: // Reservacion Proveedor
				case 28: // Nota debito compra
					break;
				default:
					btCreaAsiento.Visible = false;
					break;
			}
			switch (IdTipoFactura)
			{
				case 1: // Venta 
				case 4: // Compra
					btRetencion.Visible = true;
					break;
			}

			switch (IdTipoFactura)
			{
				case 1: // Venta 
				case 6: // Devolucion en compra
				case 4: // Compra
				case 5: // Devolucion en venta
				case 8: // Egreso de Bodega
					btGuiaRemision.Visible = true;
					break;
			}

			switch (IdTipoFactura)
			{
				case 6: // Devolucion en compra
				case 5: // Devolucion en venta
					btTransforma.Visible = false;
					break;
			}


			switch (IdTipoFactura)
			{
				case 1: // Venta
				
				
					if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'ADIGDEC'")) 				
					{
						btDigitalizacion.Visible=true;				
					}

					//	btTransforma.Visible = false;
					break;
			}
			txtBorrar.Width = 0;

			// Posicion de Botones
			int iPosDer = btBuscaArticulo.Left + btBuscaArticulo.Width;
			if (btTransforma.Visible)
			{
				btTransforma.Left = iPosDer;
				iPosDer += btTransforma.Width;
			}
			if (IdTipoFactura == 6 || IdTipoFactura == 5) // Devolucion en compra o venta
				btTransforma.Visible = true;
			if (btNotaEntrega.Visible)
			{
				btNotaEntrega.Left = iPosDer;
				iPosDer += btNotaEntrega.Width;
			}
			if (btRetencion.Visible)
			{
				btRetencion.Left = iPosDer;
				iPosDer += btRetencion.Width;
			}
			if (btCreaAsiento.Visible)
			{
				btCreaAsiento.Left = iPosDer;
				iPosDer += btCreaAsiento.Width;
			}
			if (btGuiaRemision.Visible)
			{
				btGuiaRemision.Left = iPosDer;
				iPosDer += btGuiaRemision.Width;
			}
			if (btActualizaPrecio.Visible)
			{
				btActualizaPrecio.Left = iPosDer;
				iPosDer += btActualizaPrecio.Width;
			}
			if (btNuevo.Visible)
			{
				btNuevo.Left = iPosDer;
				iPosDer += btNuevo.Width;
			}
			if (btBarras.Visible)
			{
				btBarras.Left = iPosDer;
				iPosDer += btBarras.Width;
			}
			if (btAuditoria.Visible)
			{
				btAuditoria.Left = iPosDer;
				iPosDer += btAuditoria.Width;
			}
		}
	
		private void HabilitaTab()
		{
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in utcTransaccion.Tabs)
				foreach(Control cn in ctr.TabPage.Controls)
					barraDato1.HabilitaControl(cn, barraDato1.bEditar);
			rdgContadoCred.Enabled = barraDato1.bEditar;
		}

		private void CargaProyecto()
		{
			string stExec = "Exec ProyectoPredefinido " + IdTipoFactura.ToString();
			IProyecto = Funcion.iEscalarSQL(cdsCompra, stExec + ", 1", true);
			ISubProyecto = Funcion.iEscalarSQL(cdsCompra, stExec + ", 2", true);
			IBodega = Funcion.iEscalarSQL(cdsCompra, stExec + ", 3", true);

			if (IProyecto > 0) cmbProyecto.ReadOnly = true;
			if (ISubProyecto > 0) cmbSubProyecto.ReadOnly = true;
			if (IBodega > 0) spnBodega.ReadOnly = true;
		}

		private void MoverTotales()
		{
			lblIce.Top = lblIce.Top + 16;
			lblIva.Top = lblIva.Top + 16;
			lblIva0.Top = lblIva0.Top + 16;
			lblIvaSubt.Top = lblIvaSubt.Top + 16;
			lblSubTotal.Top = lblSubTotal.Top + 16;
			lblDescuento.Top = lblDescuento.Top + 16;
			txtIceTotal.Top = txtIceTotal.Top + 16;
			txtIvaTotal.Top = txtIvaTotal.Top + 16;
			txtIva.Top = txtIva.Top + 16;
			txtIva0.Top = txtIva0.Top + 16;
			txtDescIva0Total.Top = txtDescIva0Total.Top + 16;
			txtDescIvaTotal.Top = txtDescIvaTotal.Top + 16;

			lblServicio.Visible  = false;
			lblPropina.Visible = false;
			txtServicio.Width = 0;
			txtPropina.Width = 0;
		}

		private void Compra_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				if (DialogResult.No == MessageBox.Show("Desea Cerrar esta pantalla", "Confirmación",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					e.Cancel = true;
					return;
				}
			}
			int iLong = cmbTipoCompra.Text.Trim().Length;
			if (iLong > 20) iLong = 20;

			string stAudita = "Exec AuditaCrear 53, 7, '" + cmbTipoCompra.Text.Trim().Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
		}

		private void Compra_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		#endregion load
		#region Barra Herramientas Standard
		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Buscar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
						BuscarCompra( id);
					}
					catch{}
				}
			}
		}

		public void BuscarCompra(int id)
		{
			barraDato1.IdRegistro = id;
			barraDato1.ProximoId(5);
		}
		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			#region Numero de Copias de CompraNumero
			string stCopia = "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = " 
				+ IdTipoFactura.ToString();
			int iCopias = Funcion.iEscalarSQL(cdsCompraTabla, stCopia);
			if (iCopias < 1) iCopias = 1;
			#endregion Numero de Copias de CompraNumero
			#region Muestra pantalla con totales en grande
			if (IdTipoFactura == 1 && drSeteoF["TiempoVenta"] != System.DBNull.Value
				&& (int) drSeteoF["TiempoVenta"] > 0)
			{
				double dTotal = (double) txtTotal.Value;
				double dRecibido = (double) spnVuelto.Value;
				CompraGracias miGracias = new CompraGracias((int) drSeteoF["TiempoVenta"], dTotal, dRecibido);
				miGracias.MdiParent = this.MdiParent;
				miGracias.Show();
			}
			#endregion Muestra pantalla con totales
			for (int j=0;j<iCopias;j++)
				Imprimir(iCopias);
			#region da problemas porque se queda siempre activado la impresion no se porque pero asi se arregla
			mouse_event(0x0001, 1000, 1000, 0, 0);
			mouse_event(0x0002, 10, 10, 0, 0);
			mouse_event(0x0004, 10, 10, 0, 0);
			#endregion da problemas porque se queda siempre activado no se porque
		}

		private void Imprimir(int iTipo)
		{
			Funcion.TiempoInicia();
			#region Seguridad para imprimir y reimprimir y control de vacios
			DateTime dtIni = DateTime.Now;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.ReImprimir && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			// Cuando no hay registros
			if (cmbTipoCompra.Value == System.DBNull.Value)
			{
				MessageBox.Show("No hay informacion para imprimir", "Información");
				return;
			}
			#endregion Seguridad para imprimir y reimprimir
			#region Numeracion al imprimir
			if (bNumeraImprimir && (bool)drCompra["impreso"] == false)
			{
				int iTipoNumera = 1;
				if (cmbComprobante1.ActiveRow != null && cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value != DBNull.Value)
					iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				if (iTipoNumera == 0 || IdTipoFactura == 3) // Diferente de Cotizacion
					iTipoNumera = IdTipoFactura;

				string stSelectVend = "EXEC FacturaNumeroEstacion " + iTipoNumera;

				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FNUMVEN'")&& IdTipoFactura == 1) 				
				{
					stSelectVend =  "EXEC FacturaNumeroVendedor " + drCompra["idCompra"].ToString();
				}

				string stNumero = Funcion.sEscalarSQL(cdsCompra, stSelectVend);
				string stSelect = string.Format("Update Compra Set Numero = '{0}' Where idCompra = {1}", 
					stNumero, (int) drCompra["idCompra"]);
				Funcion.EjecutaSQL(cdsSeteoF, stSelect);
				drCompra["Numero"] = stNumero;
			}
			#endregion numera al imprimir

			#region Autoimpresores Impresion
			if (bAutoImpresor && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede ReImprimir en AutoImpresores ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Autoimpresores
			
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + txtIdCompra.Value;
			if (bAutoImpresor) stReporte += ", 1";
			else stReporte += ", 0";
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			string stFiltro = "{Compra.idCompra} = " + txtIdCompra.Value.ToString();
			#endregion Nombre de Reporte

			#region Genera series en las lineas
			if (drSeteoF["SerieLinea"] != DBNull.Value && (bool) drSeteoF["SerieLinea"] == true)
			{
				string stExec = "Exec SerieLinea " + txtIdCompra.Value.ToString();
				Funcion.EjecutaSQL(cdsCompra, stExec, true);
			}
			#endregion Genera series en las lineas

			#region Imprime el reporte
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				#region Reporte con procedimiento almacenada
				string stProc = "Exec SerieExisteTipo " + txtIdCompra.Value.ToString() + ", 'FPVTISP', "
					+ cmbComprobante1.Value.ToString();
				bool bProc = Funcion.bEjecutaSQL(cdsSeteoF, stProc);
				if (bProc == true) 
				{
					stFiltro = "";
					stReporte = stReporte.Substring(0, stReporte.Length - 4) + "SP.rpt";

					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@idCompra";
					discreteVal.Value = txtIdCompra.Value.ToString();
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
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsCompra, stGrabaImprime);
			drCompra["impreso"] = true;
			barraDato1.ToolTipImprimir(Funcion.Tiempo("Imprimir"));
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtIni);
			lblTiempo.Text = ts.Seconds.ToString() + ":" + ts.Milliseconds.ToString();
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			#region Seguridad
			if (MenuLatinium.bAcademica && iCuentaTotal > 600)
			{
				MessageBox.Show("No tiene acceso a creación de Facturas en version Academica", "Confirmación", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (bAutoImpresor)
			{
				string stExec = "Exec AutoImpresorCaducado " + IdTipoFactura 
					+ ", " + cmbComprobante1.Value.ToString();
				bool bCaducado = false;
				//Funcion.bEjecutaSQL(cdsSeteoF, stExec);
				if (bCaducado)
				{
					MessageBox.Show("Caducidad de Autorizacion excede la fecha actual ",
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			#endregion Seguridad
			barraDato1.CrearRegistro();
			ValoresIniciales();
			HabilitaTab();
			#region Precio por Forma de Pago
			if ((int) drSeteoF["TipoPrecio"] == 2)
			{
				switch(IdTipoFactura)
				{
					case 1: // Factura
					case 3: // Cotizacion
						//					case 8: // Transferencia Enviada
						//					case 9: // Transferencia Recibida
						//					case 10: // Orden de Produccion
					case 13: // Reservacion Cliente
					case 11: // Pedido de Cliente
						using (CompraFormaPagoPrecio miFormaPago = new CompraFormaPagoPrecio())
						{
							miFormaPago.ShowDialog();
							drCompra["idFormaPago"] = int.Parse(miFormaPago.IdFormaPago.Value.ToString());
						}
						break;
				}
			}
			#endregion
			if (bListaCliente) cmbCliente.Select();
			else txtNombre.Select();
			ModificaGrilla(true);
			barraDato1.ToolTipNuevo(Funcion.Tiempo("Crear"));
		}

		private void ValoresIniciales()
		{
			barraDato1_Refresca(this, new System.EventArgs());
			if (IdTipoFactura == 1 || IdTipoFactura == 3) // Venta y Cotizacion
			{
				if (drSeteoF["NotaVentaP"] != DBNull.Value && (bool) drSeteoF["NotaVentaP"])
					drCompra["idComprobante"] = 2;
				else
					drCompra["idComprobante"] = 15;
				if (drSeteo["SriaAutoriz"] != System.DBNull.Value)
					drCompra["AutFactura"] = drSeteo["SriaAutoriz"];
				if (drSeteo["SriSerie"] != System.DBNull.Value)
					drCompra["SerieFactura"] = drSeteo["SriSerie"];
				if (drSeteo["SriCaduca"] != System.DBNull.Value)
					drCompra["FechaCaducidad"] = drSeteo["SriCaduca"];
				if (bAutoImpresor)
					drCompra["FechaCaducidad"] = FechaCaducidadSri;

				drCompra["idCliente"] = 0;
				if (drSeteoF["idCliente"] != DBNull.Value)
					drCompra["idCliente"] = (int) drSeteoF["idCliente"];
				if (!bListaCliente)
				{
					string stNombre = "Select Count(*) From Cliente Where idCliente = " + drCompra["idCliente"];
					int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stNombre);
					if (iCuenta > 0)
					{
						stNombre = "Select Nombre From Cliente Where idCliente = " + drCompra["idCliente"];
						drCompra["Nombre"] = Funcion.sEscalarSQL(cdsSeteoF, stNombre, false);
					}
				}
			}
			else
				drCompra["idComprobante"] = 1;

			if ((bool) drSeteoF["VentaCredito"] == true)
				drCompra["ContadoCredito"] = 2;
			else
				drCompra["ContadoCredito"] = 1;

			//Asignacion de Proyectos, Subproyectos y Bodega predefinidos o por maquina
			if (IBodega > 0) drCompra["Bodega"] = IBodega;
			if (IProyecto > 0) drCompra["idProyecto"] = IProyecto;
			if (ISubProyecto > 0) drCompra["idSubProyecto"] = ISubProyecto;

			drCompra["idFormaPago"] = 1;
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPFC'"))
				drCompra["Fecha"] = DateTime.Today;
			else
				drCompra["Fecha"] = Funcion.dtEscalarSQL(cdsSeteoF, "Select Convert(DateTime, Convert(Char(10), getdate(),112), 112)");

			drCompra["idTipoFactura"] = IdTipoFactura;
			drCompra["idSucursal"] = (int) drSeteo["NumEmpresa"];
			if (IdTipoFactura == 5 || IdTipoFactura == 6 ||
				IdTipoFactura == 27 || IdTipoFactura == 28) // Devoluciones y notas de debito
			{
				drCompra["ContadoCredito"] = 2;
				if (IdTipoFactura == 5 || IdTipoFactura == 6)
					drCompra["idComprobante"] = 4;
				else
					drCompra["idComprobante"] = 5;
			}
		} 

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			#region Verificaciones
			cmdTotal();
			if (!CambiosBotrosa())
			{
				MessageBox.Show("Ingrese Centro de Costo..... \n Campo no pude estar vacio","Información",MessageBoxButtons.OK,MessageBoxIcon.Information);
				return;
			}
			if (!Verifica()) return;
			if (!VerificaBodega()) return;
			if (!VerificaLote()) return;
			if (!VerificaVendedor()) return;
			#endregion Verificaciones
			#region Numeracion
			bool bNuevoReg = barraDato1.bNuevo;
			if (bNuevoReg && !bAutoImpresor && !bNumeraImprimir) // Numeración cuando es nuevo no es autoimpresor y no numera al imprimir
			{
				int iTipoNumera = IdTipoFactura;
				if (cmbComprobante1.ActiveRow != null && cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value != DBNull.Value)
					iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				if (iTipoNumera == 0 || IdTipoFactura == 3) // Diferente de Cotizacion
					iTipoNumera = IdTipoFactura;
				string stNumero = Funcion.sEscalarSQL(cdsCompra, "EXEC FacturaNumeroEstacion " + iTipoNumera);
				if (stNumero.Trim().Length > 0) drCompra["Numero"] = stNumero;
				else
				{
					if (txtNumero.ToString().Trim().Length == 0)
					{
						errorProvider.SetError(txtNumero, "Ingrese Codigo");
						return;
					}
				}
			}
		
			if (!bAutoImpresor && !bNumeraImprimir && !VerificaNumero()) return;
			#endregion Numeracion		
			#region Grabar y mostrar
			barraDato1.Focus();
			if (!barraDato1.GrabaRegistro())
				if (bNuevoReg)
				{
					MessageBox.Show("No se puede grabar la factura sistema ocupado", "Información");
					return;
				}

			HabilitaTab();
			ModificaGrilla(false);
			#endregion Grabar y mostrar
			#region Actualiza Bodegas a las lineas
			if ((bool) drSeteoF["VerBodega"] == false)
			{
				string ActualizaBodega = "Exec DetCompraBodegaAct " + txtIdCompra.Value;
				Funcion.EjecutaSQL(cdsSeteoF, ActualizaBodega, true);
			}
			#endregion Actualiza Bodegas a las lineas
			#region promocion
			if (bNuevoReg) // Actualiza Posicion de Nuevo Registro
			{
				barraDato1.IdRegistro = (int) txtIdCompra.Value;
				if (IdTipoFactura == 1 && drSeteoF["Promocion"] != DBNull.Value && (bool) drSeteoF["Promocion"])
				{
					string stExec = "Exec ArticuloPromocion " + txtIdCompra.Value.ToString();
					Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
					cmdTotal();
					barraDato1.GrabaRegistro();
				}
				//				if (IdTipoFactura == 4) Funcion.EjecutaSQL(cdsSeteoF, "Exec ArticuloCostosUltimos " + txtIdCompra.Value);
			}
			#endregion promocion
			#region SANTELMO
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRSTELMO'"))
			{
				Funcion.EjecutaSQL(cdsCompra, "EXEC CambiaBodega " + txtIdCompra.Value + ", " + IdTipoFactura);
			}
			#endregion SANTELMO
			#region cambia precios, consumo por receta, impresion
			if ((bool) drSeteoF["ActualizarCosto"]) 
				ActualizaPrecio(true);
			
			if ((bool) drSeteoF["ConsumoFact"]) 
				Funcion.EjecutaSQL(cdsCompra, "EXEC ConsumoFactura " + txtIdCompra.Value + ", " + IdTipoFactura);

			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPVCNSP'") && !miAcceso.ReImprimir)
				barraDato1_Impresora(this, e);
			#endregion cambia precios, consumo por receta, impresion
			barraDato1.ProximoId(5);
			barraDato1.ToolTipNuevo(Funcion.Tiempo("Grabar"));
		}

		private bool CambiosBotrosa()
		{
			bool bOk = true;
			Int32 iC = 0 ;
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRPB'")) 				
			{				
				#region Desplazamiento de grilla
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idProyecto"].Value != System.DBNull.Value)
					{						
						iC = (Int32) dr.Cells["idProyecto"].Value;						
					}
					if (iC <= 0 )
					{
						bOk=false;
					}
				}
				#endregion Desplazamiento de grilla	
			}
			return bOk;
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			string stMensajeFicha = "";
			#region Verificacion de Autorizaciones
			#region Comprobante
			string CodComprob = "";
			if (cmbComprobante1.ActiveRow != null && cmbComprobante1.ActiveRow.Cells["Codigo"].Value != null
				&& cmbComprobante1.ActiveRow.Cells["Codigo"].Value != DBNull.Value)
				CodComprob = cmbComprobante1.ActiveRow.Cells["Codigo"].Value.ToString().Trim();
			#endregion Comprobante
			if (IdTipoFactura == 4 && CodComprob != "99" &&	drSeteoF["ExigeAutorizacion"] != DBNull.Value && (bool) drSeteoF["ExigeAutorizacion"] == true)
			{
				if (drCompra["AutFactura"] == DBNull.Value || drCompra["AutFactura"].ToString().Trim().Length != 10)
				{
					errorProvider.SetError(txtAutFactura, "Autorizacion de Factura debe ser de 10 digitos");
					bOk = false;
				}
				if (drCompra["SerieFactura"] == DBNull.Value || drCompra["SerieFactura"].ToString().Trim().Length != 6)
				{
					errorProvider.SetError(txtSerieFact, "Serie de Factura debe ser de 6 digitos");
					bOk = false;
				}
				if (drCompra["FechaCaducidad"] == DBNull.Value || (DateTime) drCompra["FechaCaducidad"] < (DateTime) drCompra["Fecha"])
				{
					errorProvider.SetError(cmbCaducidad, "Fecha de Caducidad caducada");
					bOk = false;
				}
				if (!bOk)	stMensajeFicha = "Error en pestaña Autorización\n";

				string stVerificaRango = string.Format("Exec ClienteVerificaRango {0}, '{1}', {2}", 
					drCompra["idCliente"], txtNumero.Text.Trim(), cmbComprobante1.Value);
				stVerificaRango = Funcion.sEscalarSQL(cdsSeteoF, stVerificaRango);
				if (stVerificaRango.Length > 0)
				{
					errorProvider.SetError(cmbCliente, stVerificaRango);
					bOk = false;
				}
			}
			#endregion 
			#region Verifica que este en contado o credito posible con subida de datos
			if (rdgContadoCred.CheckedIndex < 0)
			{
				errorProvider.SetError(rdgContadoCred, "Ingrese forma de pago.");
				stMensajeFicha += "Error en pestaña Pago.\n";
				bOk = false;
			}
			#endregion Verifica que este en contado o credito
			#region Exige Cliente
			if (cmbCliente.Value == DBNull.Value || (int) drCompra["idCliente"] < 1)
			{
				string stMensaje = Funcion.sEscalarSQL(cdsCompra, "Exec CompraClienteVerifica " + IdTipoFactura);
				if (stMensaje.Length > 0) 
				{
					errorProvider.SetError(cmbCliente, stMensaje); //"Ingrese Nombre");
					bOk = false;
				}
			}
			#endregion Exige Cliente
			#region Limitar Credito en Ventas
			if (IdTipoFactura == 1 && Funcion.Permiso("511", cdsSeteoF) && !(bool) chkAutPrincipal.Checked)
			{
				string stExec = string.Format("Exec ClienteSaldo {0}, {1}, {2}, {3}, {4}", drCompra["idCliente"]
					, IdTipoFactura, txtIdCompra.Value, ((double)drCompra["Total"]).ToString("0.0000", us), drCompra["ContadoCredito"]);
				double dSaldoCxC = Funcion.dEscalarSQL(cdsSeteoF, stExec);
				double dMaximoCred = miProveedor.MaximoCredito;
				if (dSaldoCxC > dMaximoCred)
				{
					string stMensaje = "Infringe Cupo de Credito. Saldo: " + dSaldoCxC.ToString("#,##0.00", us) + 
						" Cupo Maximo: " + dMaximoCred.ToString("#,##0.00", us);
					errorProvider.SetError(cmbCliente, stMensaje);
					bOk = false;
				}
			}
			#endregion
			#region Exige Vendedor en Venta y Egreso de Bodega
			if ((bool) drSeteoF["ExigeVendedor"] == true && (IdTipoFactura == 1 || IdTipoFactura == 8) 
				&& (cmbVendedor.Value == DBNull.Value || (int) cmbVendedor.Value == 0))
			{
				errorProvider.SetError(cmbVendedor, "Ingrese Vendedor");
				stMensajeFicha += "Error en pestaña Nota.\n";
				bOk = false;
			}
			#endregion Exige vendedor en venta y Egreso de Bodega
			#region Verificacion de saldo de Factura Positivo y Credito si hay pagos
			double dPagos = 0;
			double dRetenido = 0;
			double dTotal = (double) drCompra["Total"];
			if (drCompra["Retenido"] != System.DBNull.Value) dRetenido = (double) drCompra["Retenido"];
			if (!barraDato1.bNuevo) // Esta editando
			{
				string stPagos = "Select Sum(IsNull(Pago, 0)) From Pago Where idCompra = " + txtIdCompra.Value;
				dPagos = Funcion.dEscalarSQL(cdsCompra, stPagos, false);
			}

			if ((int) rdgContadoCred.Value == 1 && Math.Round(dPagos, 2) > 0)
			{
				errorProvider.SetError(rdgContadoCred, "No puede modificar forma de pago. "
					+ "Lista de Pagos de esta Factura por: " + dPagos.ToString("#,##0.00"));
				stMensajeFicha += "Error en Pestaña Pago. No puede ser al contado\n";
				bOk = false;
			}
			//Calculo de Pagos reales sin pendientes
			if (!barraDato1.bNuevo)
			{
				string stPagos = "Select Sum(IsNull(Pago, 0)) From Pago Where Pendiente = 0 And idCompra = " + txtIdCompra.Value;
				dPagos = Funcion.dEscalarSQL(cdsCompra, stPagos, false);
			}
			dPagos += dRetenido;
			if ((int) rdgContadoCred.Value == 2 && Math.Round(dPagos, 2) > Math.Round(dTotal, 2) + .001) // Pago Excede factura
			{
				errorProvider.SetError(txtTotal, "Valor Total de Factura " + dTotal.ToString("#,##0.00", us)
					+ " es menor que lo cancelado " + dPagos.ToString("#,##0.00", us));
				bOk = false;
			}
			#endregion Verificacion de saldo de Factura Positivo y Credito si hay pagos
			#region Verificacion de Fecha
			if (drCompra["Fecha"] == DBNull.Value) 
			{
				errorProvider.SetError(dtFecha, "Ingrese Fecha");
				bOk = false;
			}
			// Verificacion de Periodo
			if (!miAcceso.CambiarPeriodo && drCompra["Fecha"] != DBNull.Value) 
			{
				DateTime dtFechaFact = (DateTime)drCompra["Fecha"];
				string stMensaje = string.Format("Exec CompraVerificaPeriodo '{0}'", dtFechaFact.ToString("yyyyMMdd"));
				stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stMensaje);
				if (stMensaje.Length > 0)
				{
					errorProvider.SetError(dtFecha, stMensaje);
					bOk = false;
				}
			}
			#endregion
			#region Verificacion de Centro de Costo y Proyecto en Venta y Compra
			if (IdTipoFactura == 1 || IdTipoFactura == 4)
			{
				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FVCP'"))
				{
					if ((int)cmbProyecto.Value == 0)
					{
						errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
						stMensajeFicha += "Verifique pestaña Nota. Centro de Costo\n";
						bOk = false;
					}
					if((int) cmbSubProyecto.Value == 0)
					{
						errorProvider.SetError(cmbSubProyecto, "Ingrese Proyecto");
						stMensajeFicha += "Verifique pestaña Nota. Proyecto\n";
						bOk = false;
					}
				}
			}
			#endregion
			#region Verifica Saldo Cero
			if (!miAcceso.GrabaCero && (double) txtTotal.Value == 0) 
			{
				errorProvider.SetError(txtTotal, "No puede grabar la transacción en Cero");
				bOk = false;
			}
			#endregion Verifica Saldo Cero
			#region Autoimpresor
			if (bAutoImpresor)
			{
				if (IdTipoFactura == 1 || IdTipoFactura == 3)
					if (!((int)cmbComprobante1.Value == 15 || (int) cmbComprobante1.Value == 2))
					{
						errorProvider.SetError(cmbComprobante1, "Solo se permiten Documentos Autorizados (18) o Nota de Venta en Autoimpresores");
						bOk = false;
					}
				if (IdTipoFactura == 5)
					if (txtComprobante.Text.Trim().Length==0)
					{
						errorProvider.SetError(txtComprobante, "Notas de Credito Deben tener comprobante que afectan");
						stMensajeFicha += "Verifique pestaña Nota. Comprobante\n";
						bOk = false;
					}
			}
			#endregion
			if (stMensajeFicha.Length > 0) errorProvider.SetError(utcTransaccion, stMensajeFicha.Trim());
			return bOk;
		}
		private void BorraErrores()
		{
			errorProvider.SetError(dtFecha, "");
			errorProvider.SetError(cmbCaducidad, "");
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(cmbComprobante1, "");
			errorProvider.SetError(cmbProyecto, "");
			errorProvider.SetError(cmbSubProyecto, "");
			errorProvider.SetError(cmbVendedor, "");
			errorProvider.SetError(txtAutFactura, "");
			errorProvider.SetError(txtComprobante, "");
			errorProvider.SetError(txtNumero, "");
			errorProvider.SetError(txtTotal, "");
			errorProvider.SetError(txtSerieFact, "");
			errorProvider.SetError(rdgContadoCred, "");
			errorProvider.SetError(utcTransaccion, "");
		}
		private bool VerificaLote()
		{
			if (!bLote) return true;
			if (!(IdTipoFactura == 1 || IdTipoFactura == 4  || IdTipoFactura == 5 || IdTipoFactura == 6 || IdTipoFactura == 8 
				|| IdTipoFactura == 9)) return true;
			int iLinea = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				iLinea ++;
				if (dr.Cells["idArticulo"].Value != DBNull.Value && dr.Cells["Cantidad"].Value != DBNull.Value
					&& (double)dr.Cells["Cantidad"].Value > 0)
				{
					bool bFaltaCodigo = false;
					if (dr.Cells["RefCodigo"].Value == DBNull.Value || dr.Cells["RefCodigo"].Value.ToString().Length == 0) 
						bFaltaCodigo = true;
					bool bFaltaVencimiento = false;
					if (dr.Cells["Vencimiento"].Value == DBNull.Value || ((DateTime)dr.Cells["Vencimiento"].Value).Year < 2010) 
						bFaltaVencimiento = true;
					if (bFaltaVencimiento || bFaltaCodigo)
					{
						int IdArt = (int) dr.Cells["idArticulo"].Value;
						bool bExigeLote = Funcion.bEjecutaSQL(cdsSeteoF, 
							"Select IsNull(ExigeLote, 0) From Articulo Where idArticulo = " + IdArt);
						if (!bExigeLote) continue;
						string stArticulo = "Select 'Cod: ' + Codigo + ' Item: ' + Articulo From Articulo Where idArticulo = " + IdArt;
						stArticulo = Funcion.sEscalarSQL(cdsSeteoF, stArticulo) 
							+ ".\nRequiere un lote, vencimiento y elaboración. Línea " + iLinea;
						MessageBox.Show(stArticulo, "Información");
						return false;
					}
				}
			}
			return true;
		}


		private bool VerificaVendedor()
		{
			if (!bVendedor) return true;
			int iLinea = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				
				iLinea ++;
				if (dr.Cells["idArticulo"].Value != DBNull.Value && dr.Cells["Cantidad"].Value != DBNull.Value
					&& (double)dr.Cells["Cantidad"].Value > 0)
				{
					bool bFaltaVendedor = false;
					if (dr.Cells["idVendedor"].Value == DBNull.Value || dr.Cells["idVendedor"].Value.ToString().Length == 0) 
						bFaltaVendedor = true;
					
					if (bFaltaVendedor)
					{
						int IdArt = (int) dr.Cells["idArticulo"].Value;

						string stArticulo = "Select 'Cod: ' + Codigo + ' Item: ' + Articulo From Articulo Where idArticulo = " + IdArt;
						stArticulo = Funcion.sEscalarSQL(cdsSeteoF, stArticulo) 
							+ ".\n Requiere Vendedor. Linea " + iLinea;
						MessageBox.Show(stArticulo, "Información");
						return false;
					}
				}
			}
			return true;
		}
		private bool VerificaBodega()
		{
			if (drSeteoF["SaldoPorBodega"] != DBNull.Value && (bool) drSeteoF["SaldoPorBodega"] == true)
			{
				#region Verifica que aplica revisar negativos por bodega
				// Control de Negativos con procedimientos almacenados retorna
				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPVCNSP'")) return true;

				// Verificacion de Signo Negativo
				string stVerificaSigno = "Select Signo From CompraNumero Where idTipoFactura = "
					+ IdTipoFactura.ToString();
				bool bVerificaCotiza = false;
				if (IdTipoFactura == 3 && (drSeteoF["SaldoEnCotiza"] != DBNull.Value 
					&& (bool) drSeteoF["SaldoEnCotiza"] == true)) bVerificaCotiza = true;

				// Signo no es negativo y no verifica cotizaciones retorna
				if (Funcion.iEscalarSQL(cdsSeteoF, stVerificaSigno, true) != -1 && !bVerificaCotiza)
					return true;
				#endregion

				int iBodega = (int)	drCompra["Bodega"];
				bool bBodegaLinea = false;
				if (drSeteoF["VerBodega"] != DBNull.Value && (bool) drSeteoF["VerBodega"] == true)
					bBodegaLinea = true;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != DBNull.Value && dr.Cells["Cantidad"].Value != DBNull.Value)
					{
						if (bBodegaLinea)
							if (dr.Cells["Bodega"].Value != DBNull.Value) iBodega = (int) dr.Cells["Bodega"].Value;

						double dCantidad = (double) dr.Cells["Cantidad"].Value;
						string stExec = string.Format("Exec CompraBodegaNegativo1 {0}, {1}, {2}, {3}", 
							dr.Cells["idArticulo"].Value, dCantidad.ToString("0.000000", us), iBodega, txtIdCompra.Value);

						double dSaldo = Funcion.dEscalarSQL(cdsSeteoF, stExec, true);
						if (dSaldo - dCantidad < -0.002)
						{
							int IdArt = (int) dr.Cells["idArticulo"].Value;
							string stArticulo = "Select Articulo From Articulo Where idArticulo = " + IdArt.ToString();
							stArticulo = Funcion.sEscalarSQL(cdsSeteoF, stArticulo);
							MessageBox.Show("No puede vender: " + dCantidad.ToString("0.00", us)
								+ " Saldo: " + dSaldo.ToString("0.00", us) 
								+ " Articulo: " + stArticulo, "Información");
							return false;
						}
					}
				}
			}
			return true;
		}

		private bool VerificaNumero()
		{
			bool bOk = true;

			if (txtNumero.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtNumero, "Ingrese Número de Comprobante");
				return false;
			}
			int idCompra = 0;
			if (drCompra["idCompra"] != null) idCompra = (int)drCompra["idCompra"];
			int idCliente = 0;
			if (drCompra["idCliente"] != null) idCliente = (int)drCompra["idCliente"];
			int idComprobante = (int)drCompra["idComprobante"];
			string stSelect = string.Format("Exec CompraVerificaNumero '{0}', {1}, {2}, {3}, {4}", 
				txtNumero.Text, IdTipoFactura, idComprobante, idCompra, idCliente);
			stSelect = Funcion.sEscalarSQL(cdsSeteoF, stSelect);
			if (stSelect.Length > 0)
			{
				bOk = false;
				errorProvider.SetError(txtNumero, stSelect);
			}
			return bOk;
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear ni Copiar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (MessageBox.Show("¿Desea copiar este registro?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

			Funcion.TiempoInicia();
			barraDato1.IdRegistro = CopiarFactura(IdTipoFactura);
			barraDato1.ProximoId(5);
			barraDato1.ToolTipNuevo(Funcion.Tiempo("Copiar"));
			MessageBox.Show("Copia de Factura Terminada", "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private int CopiarFactura(int idDestino)
		{
			C1.Data.C1DataRow drC = this.cdsCompra.TableViews["Compra"].AddNew();
			// Calculo de numeracion
			C1DataRow drNumero = this.cdsCompraTabla.TableViews["CompraNumero"].Rows.Find(idDestino);
			int IdSigno1 = (int) drNumero["Signo"];
			if ((bool) drNumero["Automatico"] == true)
			{
				int iTipoNumCopia = 0;
				if (cmbComprobante1.ActiveRow != null && 
					cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
					!= System.DBNull.Value)
					iTipoNumCopia = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				if (iTipoNumCopia == 0) iTipoNumCopia = idDestino;
			}// fin de numeracion
			drC["idCliente"] = drCompra["idCliente"];
			drC["Bulto"] = drCompra["Bulto"];
			drC["Peso"] = drCompra["Peso"];
			drC["Pedido"] = drCompra["Pedido"];
			drC["Depart"] = drCompra["Depart"];
			drC["Otro"] = drCompra["Otro"];
			drC["Mensaje1"] = drCompra["Mensaje1"];
			drC["Mensaje2"] = drCompra["Mensaje2"];
			drC["Mensaje3"] = drCompra["Mensaje3"];
			drC["Cif"] = drCompra["Cif"];

			drC["idComprobante"] = drCompra["idComprobante"];
			drC["idRecibe"] = drCompra["idRecibe"];
			drC["idProyecto"] = drCompra["idProyecto"];
			drC["idSubProyecto"] = drCompra["idSubProyecto"];
			drC["idFormaPago"] = drCompra["idFormaPago"];
			drC["idTipoMoneda"] = drCompra["idTipoMoneda"];
			if (idDestino == 6)	drC["idComprobante"] = 4;
			else if (idDestino == 5) drC["idComprobante"] = 4;
			else drC["idComprobante"] = drCompra["idComprobante"];
			drC["idCredTributario"] = drCompra["idCredTributario"];
			drC["ContadoCredito"] = drCompra["ContadoCredito"];
			drC["Otro"] = drCompra["Otro"];
			drC["Fecha"] = DateTime.Today;
			drC["FechaIngreso"] = DateTime.Now;
			drC["FechaVencimiento"] = DateTime.Today;
			drC["idTipoFactura"] = idDestino;
			drC["Comprobante"] = drCompra["Numero"];
			drC["Total"] = drCompra["Total"];
			drC["Desc1"] = drCompra["Desc1"];
			drC["Desc2"] = drCompra["Desc2"];
			drC["Desc3"] = drCompra["Desc3"];
			drC["Desc4"] = drCompra["Desc4"];
			drC["Descuento"] = drCompra["Descuento"];
			drC["Descuento0"] = drCompra["Descuento0"];
			drC["Iva"] = drCompra["Iva"];
			drC["SubTotalIva"] = drCompra["SubTotalIva"];
			drC["SubTotalExcento"] = drCompra["SubTotalExcento"];
			drC["Retencion"] = drCompra["Retencion"];
			drC["Ice"] = drCompra["Ice"];
			drC["Flete"] = drCompra["Flete"];
			drC["Interes"] = drCompra["Interes"];
			drC["ValorDolar"] = drCompra["ValorDolar"];
			drC["DiasCredito"] = drCompra["DiasCredito"];
			drC["FechaVencimiento"] = DateTime.Today.AddDays((int) drCompra["DiasCredito"]);
			drC["idVendedor"] = drCompra["idVendedor"];
			drC["ValorDolar"] = drCompra["ValorDolar"];
			drC["Notas"] = drCompra["Notas"];
			drC["Bodega"] = drCompra["Bodega"];
			drC["AutFactura"] = drCompra["AutFactura"];
			drC["FechaCaducidad"] = drCompra["FechaCaducidad"];
			drC["AutImprenta"] = drCompra["AutImprenta"];
			drC["SerieFactura"] = drCompra["SerieFactura"];
			// drC["Editar"] = true;
			if (!bAutoImpresor && !bNumeraImprimir)
			{
				int iTipoNumera = 0;
				if (idDestino == 1 || idDestino == 3) // Para las notas de venta
				{
					if (cmbComprobante1.ActiveRow != null && 
						cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
						!= System.DBNull.Value)
						iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				}
				if (iTipoNumera == 0) iTipoNumera = idDestino;

				string stNumeroFact = Funcion.NumeraFactura(cdsCompraTabla, iTipoNumera);
				drC["Numero"] = stNumeroFact;
			}
			cdsCompra.Update();

			// Copia de lineas
			foreach(C1.Data.C1DataRow dr in drCompra.GetChildRows())
			{
				C1.Data.C1DataRow drDC = cdsCompra.TableViews["DetCompra"].AddNew();
				drDC["idCompra"] = drC["idCompra"];
				drDC["idArticulo"] = dr["idArticulo"];
				drDC["idUnidad"] = dr["idUnidad"];
				drDC["Cantidad"] = dr["Cantidad"];
				drDC["Unidades"] = dr["Unidades"];
				drDC["Precio"] = dr["Precio"];
				drDC["Impuesto"] = dr["Impuesto"];
				drDC["Notas"] = dr["Notas"];
				drDC["RefCodigo"] = dr["RefCodigo"];
				drDC["RefNumero"] = dr["RefNumero"];
				drDC["DescuentoArt"] = dr["DescuentoArt"];
				drDC["CantidadPeso"] = dr["CantidadPeso"];
				drDC["Flete"] = dr["Flete"];
				drDC["Bodega"] = dr["Bodega"];
				drDC["DescuentoPorc"] = dr["DescuentoPorc"];
				drDC["Pedido"] = dr["Pedido"];
				drDC["Signo"] = IdSigno1;
				drDC["Codigo"] = dr["Codigo"];
				drDC["Articulo"] = dr["Articulo"];
				drDC["idProyecto"] = dr["idProyecto"];
				drDC["idSubProyecto"] = dr["idSubProyecto"];
			}
			cdsCompra.Update();

			cmdTotal();
			return (int) drC["idCompra"];
		}


		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			#region Verifica
			if (drCompra == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!miAcceso.Autorizacion && chkAutPresupuestaria.Checked)
			{
				MessageBox.Show("No puede Editar después de Autorizar", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (bAutoImpresor && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede Editar en AutoImpresores despues de imprimir ", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (drCompra["Bloquear"] != DBNull.Value && (bool) drCompra["Bloquear"] == true)
			{
				MessageBox.Show("No puede editar registro por estar bloqueado", "Seguridad",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.EditaImpresa && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede Editar Facturas Impresas ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (miAcceso.EditarSoloDia)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsSeteoF, "Select GetDate()");
				DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
				if (dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
				{
					MessageBox.Show("Solo puede editar Registros de fecha " +
						dtFechaHoy.ToString("dd/MMM/yyyy"), 
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
			HabilitaTab();
			ModificaGrilla(true);
			if (drCompra["idCliente"] != System.DBNull.Value)
				miProveedor = new Proveedor(cdsSeteoF, (int) drCompra["idCliente"]);
			if (bAutoImpresor) 
			{
				txtNumero.ReadOnly = true;
				dtFecha.ReadOnly = true;
			}
			barraDato1.ToolTipEditar(Funcion.Tiempo("Editar"));
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			BorraErrores();
			barraDato1.DeshacerRegistro();
			HabilitaTab();
			ModificaGrilla(false);
			barraDato1.ToolTipEditar(Funcion.Tiempo("Deshacer"));
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (drCompra == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if ((bool) lblAnulado.Visible == false)
			{
				//				if (bAutoImpresor)
				//				{
				//					MessageBox.Show("No puede Anular en AutoImpresores.",
				//						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//					return;
				//				}
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular " + Text.Trim(),
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				string stExecAnula = "Exec CompraBorrarVerifica " + txtIdCompra.Value + ", 0";
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExecAnula, true);
				if (stMensaje.StartsWith("No"))
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
				if (DialogResult.No == MessageBox.Show(stMensaje, 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;
		
				Funcion.TiempoInicia();
				lblAnulado.Visible = true;
				stExecAnula = "Exec CompraBorrar " + txtIdCompra.Value + ", 0";
				stMensaje = Funcion.sEscalarSQL(cdsCompra, stExecAnula, true);
				if (stMensaje.Length == 0)
				{
					barraDato1.ProximoId(5);
					barraDato1.ToolTipBorrar(Funcion.Tiempo("Anular"));
				}
				else
					MessageBox.Show(stMensaje, "Error al Anular", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			else // Si esta anulado se elimina
			{
				if (bAutoImpresor)
				{
					MessageBox.Show("No puede Eliminar en AutoImpresores.",
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (!miAcceso.Eliminar)
				{
					MessageBox.Show("No tiene acceso a Eliminar " + Text.Trim(),
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (DialogResult.No == MessageBox.Show("¿Está seguro que desea borrar el Registro?"
					+ ".\nEste proceso no puede ser revertido",	"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
					MessageBoxDefaultButton.Button2)) return;

				string stExecBorrar = "Exec CompraBorrar " + txtIdCompra.Value.ToString() + ", 1";
				string stMensaje = Funcion.sEscalarSQL(cdsCompra, stExecBorrar, true);
				if (stMensaje.Length == 0)
				{
					barraDato1.ProximoId(2);
					barraDato1.ToolTipBorrar(Funcion.Tiempo("Borrar"));
				}
				else
					MessageBox.Show(stMensaje, "Error al borrar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if ((bool) drCompra["Borrar"] == false)
			{
				MessageBox.Show("Solo puede Restaurar registros Anulados",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Recuperar)
			{
				MessageBox.Show("No tiene Acceso a Restaurar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea Restaurar el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Information)) return;

			Funcion.TiempoInicia();
			string stExec = "Exec CompraRestaurar " + txtIdCompra.Value;
			Funcion.EjecutaSQL(cdsSeteoF, stExec);
			barraDato1.ProximoId(5);
			barraDato1.ToolTipBorrar(Funcion.Tiempo("Restaurar"));
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			try
			{
				drCompra = cdsCompra.TableViews["Compra"].Rows[0];
				if ((bool)drCompra["Borrar"] == true) lblAnulado.Visible = true;
				else lblAnulado.Visible = false;
			}
			catch{};
			barraDato1.ToolTipBuscar(Funcion.Tiempo("Buscar"));
		}

		#endregion Barra Herramientas
		#region Barra Herramientas extendida
		private void btSaldo_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idArticulo"].Value == DBNull.Value)
			{
				MessageBox.Show("Seleccione el articulo que desea ver el saldo", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdArticulo = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPVCNSP'"))
			{
				CompraSaldoArt miSaldo = new CompraSaldoArt(IdArticulo);
				Cursor = Cursors.Default;
				miSaldo.ShowDialog();
				return;
			}

			string stSelect = "Exec ArticuloSaldo " + IdArticulo.ToString();
			stSelect = Funcion.sEscalarSQL(cdsSeteoF, stSelect, true);

			Cursor = Cursors.Default;
			MessageBox.Show(stSelect, "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btSaldo1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			// Muestra la forma de Entrega de un pedido
			if (IdTipoFactura == 13)
			{
				string stFiltro = "{DetCompra.RefNumero} = '" + txtNumero.Text + "'";
				Reporte miRep = new Reporte("PedidoEntrega", stFiltro);
				miRep.MdiParent = MdiParent;
				miRep.Show();
			}
		}

		private void btBuscaArticulo_Click(object sender, System.EventArgs e)
		{
			bool bExiste = true;
			try
			{
				if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value == null 
					|| ultraGrid1.ActiveRow.Cells["idArticulo"].Value == DBNull.Value)
					bExiste = false;
			}
			catch
			{
				bExiste = false;
			}

			Articulo miArticulo;
			if (bExiste)
				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'ARTMPREC'")) 				
				{
					miArticulo = new Articulo((int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value,
						(double) ultraGrid1.ActiveRow.Cells["Precio"].Value);
				}
				else				
					miArticulo = new Articulo((int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value);
			else
				miArticulo = new Articulo();

			miArticulo.ShowDialog();
			miArticulo.Dispose();
			miArticulo = null;
			Cursor = Cursors.WaitCursor;
			ArticuloCarga();
			Cursor = Cursors.Default;
		}

		private void btBuscaArticulo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
			string stTitulo = "";
			if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value != null && ultraGrid1.ActiveRow.Cells["idArticulo"].Value != DBNull.Value)
			{
				stFiltro = "{Articulo.idArticulo} = " + ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
			}
			if (drCompra["idCliente"] != DBNull.Value)
			{
				if (stFiltro.Length > 0) stFiltro += " And ";
				stFiltro += "{Cliente.idCliente} = " + drCompra["idCliente"];
				if (txtNombre.Width > 0)
					stTitulo = "Cliente " + txtNombre.Text;
				else
					stTitulo = "Cliente " + cmbCliente.ActiveRow.Cells["Nombre"].Value;
			}
			Reporte miReporte = new Reporte("Articulo por Fecha.rpt", stFiltro);
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btTransforma_Click(object sender, System.EventArgs e)
		{
			#region Verificaciones
			if (barraDato1.bEditar) return;
			if(!miAcceso.Transformacion)
			{
				MessageBox.Show("No tiene acceso a Transformar " + Text.Trim(),
					"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Verificaciones
			#region Con Procedimientos Almacenados
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTT'"))
			{
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraTransformaV " + txtIdCompra.Value);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				stMensaje = Funcion.sEscalarSQL(cdsSeteoF, 
					"Exec CompraTransformaM " + txtIdCompra.Value.ToString() + ", 1", true);
				if (DialogResult.Yes == MessageBox.Show(stMensaje, 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{
					this.Cursor = Cursors.WaitCursor;
					int iDestino = Funcion.iEscalarSQL(cdsSeteoF, 
						"Exec CompraTransformaM " + txtIdCompra.Value.ToString() + ", 2", true);
					int idCompraDest = Funcion.iEscalarSQL(cdsSeteoF, 
						"Exec CompraTransformaCopia " + txtIdCompra.Value.ToString(), true);
					Compra miCompra1 = new Compra(iDestino, idCompraDest);
					miCompra1.MdiParent = this.MdiParent;
					miCompra1.Show();
					this.Cursor = Cursors.Default;
				}
				return;
			}
			#endregion Con Procedimientos Almacenados
			#region Verificacion 2
			if ((bool) lblAnulado.Visible == true)
			{
				MessageBox.Show("No puede transformar registros Anulados ",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			#endregion Verificacion 2
			#region Nombre Destino y Condiciones
			string stTransforma = "";
			int idDestino = 0;
			if (drCompra["idCliente"] != DBNull.Value)
				miProveedor = new Proveedor(cdsSeteoF, (int) drCompra["idCliente"]);
			switch(this.IdTipoFactura)
			{
				case (int) Funcion.TipoFactura.Venta:
					stTransforma = "Desea crear una devolución en venta de esta Venta";
					idDestino = (int) Funcion.TipoFactura.DevolucionVenta;
					break;
				case (int) Funcion.TipoFactura.OrdenCompra:
					#region Revisar no tiene sentido aparentemente
					//					string stUsuarioT = "Exec VerificaUsuarioT " + txtIdCompra.Value.ToString();
					//					int iVerifica = Funcion.iEscalarSQL(cdsSeteoF, stUsuarioT, true);
					//					if (iVerifica == 0) 
					//					{
					//						MessageBox.Show("No tiene el cupo de credito para transformar", "Información");
					//						return;
					//					}
					#endregion Revisar no tiene sentido aparentemente

					#region Transforma con Lista de opciones
					if (drSeteoF["OrdenCIngreso"] != DBNull.Value && (bool) drSeteoF["OrdenCIngreso"] == true)
					{
						string stSelect = "Select Count(*) From CompraNumero Where IsNull(Transforma, 0) = 2";
						int iCuenta1 = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
						if (iCuenta1 > 0)
						{
							ListaTransformacion miLista = new ListaTransformacion();
							if (DialogResult.Cancel == miLista.ShowDialog())
							{
								miLista.Dispose();
								return;
							}
							stTransforma = "Desea crear un " + miLista.txtNombre.Text + " con esta orden de compra";
							idDestino = (int) miLista.txtIdFactura.Value;
							miLista.Dispose();
						}
						else
						{
							MessageBox.Show("Defina las transformaciones en Archivo - Propiedades - Facturacion - Numeracion "
								+ "\n ingresando en la Columna Transforma el valor 2 en las filas a transformar", "Información"
								, MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
					}
						#endregion Transforma con Lista de opciones
						#region Transforma a Pedido Proveedor
					else if (drSeteoF["OrdenCPedido"] != DBNull.Value && (bool) drSeteoF["OrdenCPedido"] == true)
					{
						stTransforma = "Desea crear un Pedido de Proveedor con esta Orden de Compra";
						idDestino = (int) Funcion.TipoFactura.PedidoProveedor;
					}
						#endregion Transforma a Pedido Proveedor
						#region Transforma a Factura
					else
					{
						stTransforma = "Desea crear una Factura de Compra con esta orden de compra";
						idDestino = (int) Funcion.TipoFactura.Compra;
					}
					break;
					#endregion Transforma a Factura
				case (int) Funcion.TipoFactura.Cotizacion:
					#region Transforma a Egreso si tiene autorizacion Principal
					if (drSeteoF["CotizaEgreso"] != DBNull.Value &&	(bool) drSeteoF["CotizaEgreso"] == true)
					{
						if (!chkAutPrincipal.Checked)
						{
							MessageBox.Show("No puede transformar una cotización a Egreso sin Autorización Principal", 
								"Información");
							return;
						}
						stTransforma = "Desea crear un Egreso con esta Cotización";
						idDestino = (int) Funcion.TipoFactura.EgresoBodega;
					}
						#endregion Transforma a Egreso si tiene autorizacion Principal
						#region Transforma a Pedido de Cliente
					else if (drSeteoF["CotizaReserva"] == DBNull.Value || (bool) drSeteoF["CotizaReserva"] == true)
					{
						stTransforma = "Desea crear un Pedido de Cliente con esta Cotización";
						idDestino = (int) Funcion.TipoFactura.PedidoCliente;
					}
						#endregion Transforma a Pedido de Cliente
						#region Transforma a Venta
					else
					{
						stTransforma = "Desea crear una Factura de Venta con esta Cotización";
						idDestino = (int) Funcion.TipoFactura.Venta;
					}
					break;
					#endregion Transforma a Venta
				case (int) Funcion.TipoFactura.Compra:
					stTransforma = "Desea crear una Devolución de compra con esta Compra";
					idDestino = (int) Funcion.TipoFactura.DevolucionCompra;
					break;
				case (int) Funcion.TipoFactura.DevolucionVenta: // No implementada
					stTransforma = "Desea Abonar esta devolucion a su cuenta por Cobrar";
					return;
				case (int) Funcion.TipoFactura.EgresoBodega:
					stTransforma = "Desea crear un Ingreso a Bodega con este Egreso de Bodega";
					idDestino = (int) Funcion.TipoFactura.IngresoBodega;
					break;
				case (int) Funcion.TipoFactura.IngresoBodega:
					stTransforma = "Desea crear un Egreso de Bodega con este Ingreso a bodega";
					idDestino = (int) Funcion.TipoFactura.EgresoBodega;
					break;
				case (int) Funcion.TipoFactura.ReservaCliente:
					stTransforma = "Desea crear una Factura de Venta con esta Reservacion del Cliente";
					idDestino = (int) Funcion.TipoFactura.Venta;
					break;
				case (int) Funcion.TipoFactura.ProdTerminado:
					stTransforma = "Desea crear un Egreso de Bodega de los componentes del Producto Terminado";
					idDestino = (int) Funcion.TipoFactura.EgresoBodega;
					break;
				case (int) Funcion.TipoFactura.ReservaProveedor:
					stTransforma = "Desea crear una Factura de Compra con esta Reservacion del Proveedor";
					idDestino = (int) Funcion.TipoFactura.Compra;
					break;
				case (int) Funcion.TipoFactura.PedidoCliente:
					stTransforma = "Desea crear una Factura de Venta con este Pedido del cliente";
					idDestino = (int) Funcion.TipoFactura.Venta;
					break;
				case (int) Funcion.TipoFactura.PedidoProveedor:
					stTransforma = "Desea crear un Ingreso de Bodega con este Pedido de proveedor";
					idDestino = (int) Funcion.TipoFactura.IngresoBodega;
					break;
				case (int) Funcion.TipoFactura.ImpTransito:
					stTransforma = "Desea crear un Ingreso de Bodega con esta Importación en Transito";
					idDestino = (int) Funcion.TipoFactura.IngresoBodega;
					break;
				default:
					MessageBox.Show("Transformación no implementada");
					return;
			}
			#endregion Nombre Destino y Condiciones
			#region Confirmacion y ejecucion
			if (DialogResult.Yes == MessageBox.Show(stTransforma, "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				#region Producto Terminado
				if (IdTipoFactura == (int) Funcion.TipoFactura.ProdTerminado)
				{
					string stExecPT = "Exec ConsumoFactura " + txtIdCompra.Value + ", 12";
					string stNegativoReceta = Funcion.sEscalarSQL(cdsSeteoF, stExecPT, true);
					if (stNegativoReceta.Length > 0) 
					{
						MessageBox.Show(stNegativoReceta, "Información");
						return;
					}
					MessageBox.Show("Egreso de Bodega Generado", "Información");
					return;
				}
				#endregion Producto Terminado
				#region Cotizacion
				if (IdTipoFactura == (int) Funcion.TipoFactura.Cotizacion && idDestino != (int) Funcion.TipoFactura.EgresoBodega)
				{
					string stExec1 = string.Format("Exec ClienteSaldo {0}, {1}, {2}, {3}, {4}", drCompra["idCliente"]
						, IdTipoFactura, txtIdCompra.Value, ((double)drCompra["Total"]).ToString("0.0000", us), drCompra["ContadoCredito"]);
					double dSaldoCxC = Funcion.dEscalarSQL(cdsSeteoF, stExec1);
					double dMaximoCred = miProveedor.MaximoCredito;
					if (dSaldoCxC > dMaximoCred)
					{
						string stMensaje = "Infringe Cupo de Credito. Saldo: " + dSaldoCxC.ToString("#,##0.00", us) + 
							" Cupo Maximo: " + dMaximoCred.ToString("#,##0.00", us);
						errorProvider.SetError(cmbCliente, stMensaje);
						return;
					}

					if (!(bool) drSeteoF["SaldoNegativo"] || ListaNegativos() == 0)
					{
						Cursor = Cursors.WaitCursor;
						string stExec = "EXEC TransformaCotizacion " + txtIdCompra.Value;
						Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
						Compra miCompra = new Compra(idDestino);
						miCompra.MdiParent = MdiParent;
						miCompra.Show();
						Cursor = Cursors.Default;
						return;
					}
					return;
				}
				#endregion Cotizacion
				#region Reserva de Cliente y de Proveedor
				if (IdTipoFactura == (int) Funcion.TipoFactura.ReservaCliente || IdTipoFactura == (int) Funcion.TipoFactura.ReservaProveedor)
				{
					string stNum = Funcion.NumeraFactura(cdsCompra, idDestino);
					if (stNum.Trim().Length > 0) drCompra["Numero"] = stNum;
					drCompra["idTipoFactura"] = idDestino;
					cdsCompra.Update();
					IdTipoFactura = idDestino;
					int iSigno = 1;
					if (idDestino == (int) Funcion.TipoFactura.Venta) iSigno = -1;
					stNum = "Update detCompra Set Signo = " + iSigno + " Where idCompra = " + txtIdCompra.Value;
					Funcion.EjecutaSQL(cdsCompra, stNum, true);
					MessageBox.Show("Reservacion Transformada en factura", "Información");
					Compra2_Load(this, e);
					return;
				}
				#endregion Reserva de Cliente y de Proveedor
				#region otros Casos
				Cursor = Cursors.WaitCursor;
				int IdOrigen = (int) txtIdCompra.Value;
				int idC = CopiarFactura(idDestino);
				#endregion otros Casos
				#region Destino Ingreso de Bodega
				if (idDestino == (int) Funcion.TipoFactura.IngresoBodega) // Copia serie de Egreso
				{
					string stCopiar = "Exec SerieCopiar " + IdOrigen.ToString() + ", " + idC.ToString();
					Funcion.EjecutaSQL(cdsSeteoF, stCopiar, true);
				}
				#endregion Destino Ingreso de Bodega
				#region SANTELMO
				if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRSTELMO'"))
				{
					Funcion.EjecutaSQL(cdsCompra, "EXEC CambiaBodega " 
						+ idC.ToString()  + ", " + idDestino.ToString());
				}
				#endregion SANTELMO
				#endregion Confirmacion y ejecucion

				#region Abre la pantalla de destino
				Compra miCompra1 = new Compra(idDestino, idC);
				miCompra1.MdiParent = this.MdiParent;
				miCompra1.Show();
				Cursor = Cursors.Default;
				#endregion Abre la pantalla de destino
			}
		}

		private int ListaNegativos()
		{
			int num3 = 0;
			string text1 = "";
			foreach (UltraGridRow row1 in this.ultraGrid1.Rows.All)
			{
				int num1 = 0;
				double num2 = 0;
				if (row1.Cells["idArticulo"].Value != DBNull.Value)
				{
					num1 = (int) row1.Cells["idArticulo"].Value;
				}
				if (row1.Cells["Cantidad"].Value != DBNull.Value)
				{
					num2 = (double) row1.Cells["Cantidad"].Value;
				}
				if (num1 > 0)
				{
					string text2 = "Select IsNull(Saldo, 0) From Articulo Where idArticulo = " + num1.ToString();
					double num4 = Funcion.dEscalarSQL(this.cdsSeteoF, text2, true);
					if (num4 < num2)
					{
						num3++;
						text2 = "Select Articulo From Articulo Where idArticulo = " + num1.ToString();
						text1 = text1 + Funcion.sEscalarSQL(cdsSeteoF, text2, true);
						string text3 = text1;
						text1 = text3 + " Saldo " + num4.ToString("0.00", us) + " Pedido " + num2.ToString("0.00", this.us) + "\n";
					}
				}
			}
			if (num3 > 0)
			{
				MessageBox.Show(text1, "Información");
			}
			return num3;
		}

		private void btTransforma_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar) return;
			if (ultraGrid1.ActiveRow == null) return;
			int IdDetCompra = (int) ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
			using (AsientoDist miDist = new AsientoDist(IdDetCompra, true))
			{
				try
				{
					miDist.ShowDialog();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
				barraDato1.ProximoId(5);
			}
		}

		private void btNotaEntrega_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;

			ImprimeNota(0);
		}

		private void ImprimeNota(int iTipo)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Nota de Entrega", "Seguridad",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idCompra} = " + txtIdCompra.Value;
			string stReporte = "NotaEntrega.rpt";
			if (IdTipoFactura == (int) Funcion.TipoFactura.Compra) 
				stReporte = "IngresoBodega.rpt";
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			if (iTipo == 0)
			{
				miReporte.MdiParent = MdiParent;
				miReporte.Show();
			}
			else
				miReporte.ImprimeRep();

			Cursor = Cursors.Default;
		}

		private void btNotaEntrega_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			#region Verificaciones
			if (e.Button != MouseButtons.Right)	return;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir o Enviar por Email" + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea enviar estos datos por Email?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			#endregion Verificaciones
			#region Verificacion de Configuracion de Correo la Empresa
			string stCorreoOrigen = drSeteo["EMail"].ToString();
			if (stCorreoOrigen.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese EMail, en Archivo - Propiedades - Generales Pestaña: Básico./nEjemplo: infoelect@yahoo.es", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string stClaveCorreo = drSeteo["ClaveCorreo"].ToString();
			if (stClaveCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Contraseña, en Archivo - Propiedades - Generales Pestaña: Correo Electrónico", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stUsuarioCorreo = drSeteo["UsuarioCorreo"].ToString();
			if (stUsuarioCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Usuario, en Archivo Propiedades Generales en la Pestaña con Correo Electronico./nEjemplo: infoelect", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stPuertoCorreo = drSeteo["PuertoCorreo"].ToString();
			if (stPuertoCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Puerto de Correo, en Archivo Propiedades Generales en la Pestaña con Correo Electronico. Ej. 465", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (drSeteo["SslCorreo"] == DBNull.Value)
			{
				MessageBox.Show("Ingrese habilitacion ssl del Correo, en Archivo Propiedades Generales en la Pestaña con Correo Electronico", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			bool bEnabledSSL = bool.Parse(drSeteo["SslCorreo"].ToString());

			string stServidorCorreo = drSeteo["ServidorCorreo"].ToString();
			if (stServidorCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Servidor de Correo Saliente Ej. smtp.correo.yahoo.es, en Archivo Propiedades Generales Complemento", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Verificacion de Configuracion de Correo la Empresa
			#region Verificacion de Correo electronico del cliente
			if (cmbCliente.Value == DBNull.Value || (int) cmbCliente.Value == 0)
			{
				MessageBox.Show("Ingrese Cliente", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stDestino = "Select IsNull(EMail, '') From Cliente Where idCliente = "	+ cmbCliente.Value;
			stDestino = Funcion.sEscalarSQL(cdsSeteoF, stDestino, true);
			if (stDestino.Trim().Length == 0)
			{
				MessageBox.Show("Cliente no tiene correo electronico", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Verificacion de Correo electronico del cliente
			Imprimir(-1); // Genera Archivo pdf
			#region Configuracion de datos de Email
			MailMessage msg = new MailMessage();
			msg.To = stDestino;
			msg.From = stCorreoOrigen;
			msg.Subject = cmbTipoCompra.Text.Trim() + " " + txtNumero.Text.Trim();
			msg.Body = "Revisar informacion de archivo Adjunto";
			msg.Priority = MailPriority.Normal;
			msg.BodyFormat = MailFormat.Text;			//o MailFormat.Html
			#endregion Configuracion de Email
			#region Activar Autenticación
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1"); 
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", stUsuarioCorreo);// Nombre de usuario
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", stClaveCorreo);  // Contraseña
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", stPuertoCorreo);
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", bEnabledSSL);
			#endregion Activar Autenticación
			#region Adjunta archivo y envia
			SmtpMail.SmtpServer = stServidorCorreo; 	// El servidor de correo
			try
			{
				MailAttachment Attachment = new MailAttachment(MenuLatinium.stDirInicio + @"\Correo.pdf");
				//				MailAttachment Attachment = new MailAttachment("Correo.pdf");
				msg.Attachments.Add(Attachment);
				SmtpMail.Send(msg); 
				MessageBox.Show("Correo Electronico Enviado", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "error en envio de mail",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion Adjunta archivo y envia
		}


		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede crear una retención mientras está en edición"
					, "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Retencion)
			{
				MessageBox.Show("No tiene acceso a creación de retenciones",
					"Seguridad",MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Cursor = Cursors.WaitCursor;
			if ((int) cmbTipoCompra.Value == 4) // Compra
			{
				using (EgresoAT miEgreso = new EgresoAT())
				{
					miEgreso.IdFactura = (int) this.txtIdCompra.Value;
					miEgreso.drFactura = this.drCompra;
					miEgreso.ShowDialog();
				}
				barraDato1.ProximoId(5);
			}

			if ((int) this.cmbTipoCompra.Value == 1) // Venta
			{
				using(IngresoAT miIngreso = new IngresoAT())
				{
					miIngreso.IdFactura = (int) txtIdCompra.Value;
					miIngreso.drFactura = drCompra;
					miIngreso.ShowDialog();
				}
				
				barraDato1.ProximoId(5);
			}
			Cursor = Cursors.Default;
		}

		private void btRetencion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			#region Verificaciones
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede crear el asiento de la retención mientras está en edición"
					, "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if ((double) drCompra["Retenido"] < .001)
			{
				MessageBox.Show("No existe retención", "Información");
				return;
			}
			#endregion Verificaciones
			#region Confirmacion
			if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento de Retencion?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			#endregion Confirmacion

			#region Verifica si Existe y confirma sobreescribir
			int IdAsiRet = 0;
			if (drCompra["idAsiRet"] != DBNull.Value && (int) drCompra["idAsiRet"] > 0)
			{
				IdAsiRet = (int) drCompra["idAsiRet"];
				string stSelect = "SELECT COUNT(*) FROM Asiento WHERE Borrar = 0 And idAsiento = "+
					IdAsiRet.ToString().Trim();
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect);

				if (iCuenta > 0)
				{
					if (DialogResult.No == MessageBox.Show("Asiento de Retencion ya existe. ¿Desea sobreescribirlo?", 
						"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
					{
						Cursor = Cursors.Default;
						return;
					}
					// Si desea sobrrescribir
					string stBorrar = "DELETE FROM AsientoDetalle WHERE idAsiento = "+
						IdAsiRet.ToString().Trim();
					//					bRepetido = true;
					Funcion.EjecutaSQL(cdsSeteoF, stBorrar);
				}
			}
			#endregion Verifica si Existe y confirma sobreescribir

			#region Genera Asiento solo de retencion
			Cursor = Cursors.WaitCursor;
			int IdCompra = int.Parse(txtIdCompra.Value.ToString());

			string stExec = "EXEC AsientoRetencionC " + IdAsiRet.ToString() + ", -"
				+ IdCompra.ToString() + ", " + IdTipoFactura.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
			MessageBox.Show("Asiento de Retencion Generado", "Información");
			Cursor = Cursors.Default;
			#endregion Genera Asiento solo de retencion
		}

		private void btCreaAsiento_Click(object sender, System.EventArgs e)
		{
			#region Verificaciones
			// Crea asiento de Factura
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
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
			string stProcedimiento = Funcion.sEscalarSQL(cdsSeteoF, 
				"Exec CompraAsientoTipo " + txtIdCompra.Value, true);

			if (stProcedimiento.Length == 0)
			{
				MessageBox.Show("Tipo de Transaccion no genera Asiento", "Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Verifica si la transaccion genera Asiento
			#region Confirmacion
			if (btCreaAsiento.Visible)
				if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento?", "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					return;

			this.Cursor = Cursors.WaitCursor;
			#endregion Confirmacion
			#region verificacion de Existencia
			string stExisteAsiento = "Exec AsientoMaestroExiste "	+ txtIdCompra.Value;
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				Cursor = Cursors.Default;
				return;
			}
			stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 2", true);
			if (stMensaje.Length > 0)
			{
				if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
				{
					this.Cursor = Cursors.Default;
					return;
				}
				stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExisteAsiento + ", 3", true);
			}
			#endregion verificacion
			#region Crea Asiento
			string stGeneraCompra = "EXEC " + stProcedimiento + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stGeneraCompra, true);
			MessageBox.Show("Asiento Generado con éxito", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			#endregion Crea Asiento
		}

		private void btCreaAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (txtIdAsiento.Value == DBNull.Value || (int) txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe un asiento creado de esta transacción");
				return;
			}
			int IdAsiento = (int) txtIdAsiento.Value;
			if (drCompra["idAsientoPer"] != DBNull.Value && (int) drCompra["idAsientoPer"] > 0)
			{
				MessageBox.Show("Origen es un Asiento de Comprobante", "Información");
				IdAsiento = (int) drCompra["idAsientoPer"];
			}
			
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void btGuiaRemision_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			CompraGuia miGuia = new CompraGuia((int) txtIdCompra.Value);

			miGuia.dtFecha = (DateTime) dtFecha.Value;
			int idCliente = 0;
			if (cmbCliente.Text.Trim().Length > 0) idCliente = (int) cmbCliente.Value;
			miGuia.IdCliente = idCliente;
			miGuia.MdiParent = MdiParent;
			miGuia.Show();
		}

		private void btGuiaRemision_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idCompra} = " + txtIdCompra.Value;
			using (Reporte miReporte = new Reporte("GuiaRemision.Rpt", stFiltro))
			{
				miReporte.ShowDialog();
			}
			Cursor = Cursors.Default;
		}


		private void btActualizaPrecio_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			if (IdTipoFactura != 4)
			{
				MessageBox.Show("Actualización de precios solo en compras");
				return;
			}
			if (!miAcceso.ActualizaPrecio)
			{
				MessageBox.Show("No tiene acceso a Actualización de precios");
				return;
			}
			ActualizaPrecio(false);
		}

		private void ActualizaPrecio(bool Verifica)
		{
			if (IdTipoFactura != 4) return;
			if (barraDato1.bEditar) return;
			if (!Verifica)
			{
				if (DialogResult.No == MessageBox.Show(
					"¿Desea Actualizar precios en base a costo de Compra + % de Incremento?",
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					return;
			}
			string stVerifica = ", 0";
			if (Verifica) stVerifica = ", 1";

			string stExec = "Exec ArticuloActualizaPrecio " + txtIdCompra.Value + stVerifica;
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);

			if (btActualizaPrecio.Visible)
			{
				MessageBox.Show("Precios Actualizados",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void btActualizaPrecio_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!(IdTipoFactura == (int) Funcion.TipoFactura.IngresoBodega || 
				IdTipoFactura == (int) Funcion.TipoFactura.EgresoBodega)) return;
			if (DialogResult.No == MessageBox.Show("Desea Actualizar precio a costo promedio",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stExec = "Exec CompraPrecioCosto " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
			barraDato1.ProximoId(5);
			MessageBox.Show("Precios actualizados al costo promedio", "Información");
		}
		private void btNuevo_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			this.ultraGrid1.Rows.Band.AddNew();
			if (!(bool) drSeteoF["VerDesplegable"])
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Codigo"];
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			}
			else
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			}
		}

		private void btBarras_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stExec = "Exec BarraFactura " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
			Reporte miRep = new Reporte("BarrasArticulo.rpt", "{BarrasCodigos.idArticulo} > 0");
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;		
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(IdTipoFactura, (int) txtIdCompra.Value, true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(IdTipoFactura, (int) txtIdCompra.Value, false);
			Cursor = Cursors.Default;
		}

		#endregion Barra Herramientas extendida
		#region cds
		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			cdsArticulo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			//
			//			string stFiltro = "Exec ArticuloCargarTransaccion " + IdTipoFactura;
			//			stFiltro = Funcion.sEscalarSQL(cdsSeteoF, stFiltro, true);
			//			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsArticulos"];
			//			stFiltro = "1=0";
			//			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Articulo"], stFiltro));
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
			//
			//			string stFiltro = "Exec ClienteCarga " + IdTipoFactura;
			//			stFiltro = Funcion.sEscalarSQL(cdsCliente, stFiltro, true);
			//			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];
			//
			//			stFiltro = "1=0";
			//			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"], stFiltro));
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompra_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_Compra")
			{
				drCompra = e.Row;
				lblAnulado.Visible = (bool) drCompra["Borrar"];
				if ((int) e.Row["idTipoFactura"] < 1)	
					e.Row["idTipoFactura"] = IdTipoFactura;
			}
		}

		private void cdsComprobante_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}


		private void cdsTipoImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoImportacion.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
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

		private void cdsUnidad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUnidad.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		#endregion cds
		#region Ultragrid
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region el siguiente codigo es para filtrar cc por proy
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRPBP'")) 				
			{
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "CentroCostoBR")
					{
						if(controlW != e.Cell.Row.Cells["CentroCostoBR"].Value.ToString())
						{
							try
							{
								controlCCP();
								controlW = e.Cell.Row.Cells["CentroCostoBR"].Value.ToString();							
								NuevoCodigo = codCC(controlW);
								NuevoCodigo=NuevoCodigo.Substring(0,5);	
								if (NuevoCodigo.Length >0)
								{
									string stSelect = "Select idProyecto From Proyecto Where CentroCosto like '"
										+ NuevoCodigo + "%'";
									int IdProy = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
									e.Cell.Row.Cells["idProyecto"].Value = IdProy;
								}
							}
							catch{};

						}
						//						else return;
					}
				}
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "ProyectoBR")
					{
						try
						{
							if(ultraGrid1.Rows.Count < 1 )return;
							string CodCentroCosto = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();							
							string NuevoCodigo1 = codCC(CodCentroCosto);
							NuevoCodigo1=NuevoCodigo1.Substring(0,5);	

							if(controlW1 != e.Cell.Row.Cells["ProyectoBR"].Value.ToString())
							{
								controlCPP(NuevoCodigo1);							
								controlW1 = e.Cell.Row.Cells["ProyectoBR"].Value.ToString();
								if (controlW1.Trim().Length >0)
								{

									string stSelect1 = "Select idSubProyecto From subProyecto Where nombre like '"
										+ controlW1 + "%'";
									int IdProy1 = Funcion.iEscalarSQL(cdsSeteoF, stSelect1, true);
									e.Cell.Row.Cells["idSubProyecto"].Value = IdProy1;
								}

							}
						}
						catch{}
					}	
					//					else return;
					
				}
			}
			#endregion el siguiente codigo es para filtrar cc por proy
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
				int IdArt = 0;
				string stCod = "";
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
				string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}",
					stCod, IdTipoFactura);
				if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}",
												e.Cell.Column.ToString(), stCod));
				iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				#endregion Cuenta coincidencias del codigo
				#region Mas de una coincidencia o ninguna muestra la lista
				if (iCuenta != 1)
				{
					if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}",
													e.Cell.Column.ToString(), stCod));

					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, IdTipoFactura))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							if (bRastreo) Rastreo("Articulo escogido de la lista");
							IdArt = miBusqueda.idArticulo;
						}
						else
						{
							if (bRastreo) Rastreo("Cancela la seleccion de Articulos en lista");
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
					stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, '', '', 0, 0, 0"
						, stCod, IdTipoFactura);
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
			#region Unidad
			if (e.Cell.Column.ToString() == "Unidad")
			{
				MessageBox.Show("Cambio en Unidad!");
				if ((int) drSeteoF["TipoPrecio"] == 1 || (int) drSeteoF["TipoPrecio"] == 3) Precio(e);
			}
			#endregion Unidad
			#region Verificaciones de cambio de articulo
			if (VerificaCambio > 0)
			{
				#region Limites
				if (!Limites(e)) 
				{
					if (bRastreo) Rastreo("No cumple limites");
					bActualiza = true;
					return;
				}
				#endregion Limites
				#region Precio
				if (VerificaCambio == 2) // Cambia de Precio
					Precio(e);
				else
				{
					// Cambio de precio por cantidad o por procedimiento
					if ((int) drSeteoF["TipoPrecio"] == 1 || (int) drSeteoF["TipoPrecio"] == 3) 
						Precio(e);
				}
				#endregion Precio
				DescuentoPorcentaje(e);
				Promocion(e);
				cmdTotal();
			}
			#endregion Verificaciones
			#endregion Fin de Verificacion de Cambio
			#region DescuentoPorc o DescuentoArt
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				bActualiza = false;
				DescuentoPorcentaje(e);
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

		private void CambiaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((bool) drSeteoF["VerUnidad"] == false) return;
			int IdUnidad = (int) e.Cell.Row.Cells["idUnidad"].Value;
			int IdArticulo = (int) e.Cell.Row.Cells["idArticulo"].Value;
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
			}
			miReader.Close();
		}

		private bool Limites(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Verifica valores nulos de cantidad o de idArt
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			#endregion Verifica valores nulos de cantidad o de idArt

			#region Verifica si aplica revision
			string stIdArt = e.Cell.Row.Cells["idArticulo"].Value.ToString();
			double dCant = (double) e.Cell.Row.Cells["Cantidad"].Value;
			string stVerif = "Exec DetCompraVerifNegValida " + stIdArt + ", "
				+ dCant.ToString("0.000000", us) + ", " + IdTipoFactura;
			if(!Funcion.bEjecutaSQL(cdsSeteoF, stVerif)) return true;
			#endregion verifica si aplica revision

			#region Verifica Negativos o alertas
			string stBodega = cmbBodega.Value.ToString();
			if ((bool) drSeteoF["VerBodega"])
				stBodega = e.Cell.Row.Cells["Bodega"].Value.ToString();
			string stidDetC = e.Cell.Row.Cells["idDetCompra"].Value.ToString();
			stVerif = "Exec DetCompraVerifNegativo " + stIdArt + ", "
				+ dCant.ToString("0.000000", us) + ", " + stBodega + ", " + stidDetC + ", " + IdTipoFactura;
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stVerif, true);
			#endregion Verifica Negativos o alertas

			#region Mensaje de Negativo (empieza con -) o Alerta
			if (stMensaje.Length > 0)
			{
				if (stMensaje.StartsWith("-")) // En negativo cambia cantidad a 0
					e.Cell.Row.Cells["Cantidad"].Value = 0;
				MessageBox.Show(stMensaje.Substring(1), "Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return false;
			}
			#endregion Mensaje de Negativo (empieza con -)o Alerta
			return true;
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
			int idFormaPago = (int) cmbFormaPago.Value;
			#endregion Variables
			#region Lee Precio
			string stExec = string.Format("Exec ArticuloPrecio {0}, {1}, {2}, {3}, {4}",
				IdArt, idCliente, dCantidad.ToString("0.000000", us), idFormaPago, IdTipoFactura);
			double fPrecio = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			if (fPrecio >=0) e.Cell.Row.Cells["Precio"].Value = fPrecio;
			#endregion Lee Precio
			#region Lee Iva
			stExec = string.Format("Exec ArticuloIva {0}, {1}",
				IdTipoFactura, IdArt);
			double fIva = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			e.Cell.Row.Cells["Impuesto"].Value = fIva;
			#endregion Lee Iva
			#region Borrar
			//			miArt = new Items(cdsSeteoF, IdArt);
			//			e.Cell.Row.Cells["Impuesto"].Value = miArt.Iva;
			//			if (!(IdTipoFactura == (int) Funcion.TipoFactura.IngresoBodega || IdTipoFactura == (int) Funcion.TipoFactura.EgresoBodega))
			//				if ((int) drSeteoF["TipoPrecio"] == 4) return; // Precios manuales
			//
			//			double fPrecio = 0;
			//			// Calculo de precio
			//			switch (IdTipoFactura)
			//			{
			//				case (int)Funcion.TipoFactura.IngresoBodega:
			//					e.Cell.Row.Cells["Impuesto"].Value = 0;
			//					return;
			//				case (int)Funcion.TipoFactura.Venta:
			//				case (int)Funcion.TipoFactura.Cotizacion:
			//				case (int)Funcion.TipoFactura.EgresoBodega:
			//					//				case 10: // Orden de Produccion
			//				case (int)Funcion.TipoFactura.ReservaCliente:
			//				case (int)Funcion.TipoFactura.PedidoCliente:
			//					if (IdTipoFactura == (int)Funcion.TipoFactura.EgresoBodega || IdTipoFactura == 9)
			//						if (drSeteoF["PrecioTE"] != DBNull.Value && !(bool) drSeteoF["PrecioTE"])
			//						{
			//							fPrecio = miArt.CostoPromedio;
			//							e.Cell.Row.Cells["Precio"].Value = Math.Round(fPrecio, 5);
			//							e.Cell.Row.Cells["Impuesto"].Value = 0;
			//							return;
			//						}
			//				switch ((int) drSeteoF["TipoPrecio"])
			//				{
			//						#region Precio por Cliente
			//					case 0: // Precio definido por cliente
			//						int stPrecio = 1;
			//						stPrecio = (int) miProveedor.idTipoPrecio;
			//						fPrecio = (double) miArt.Precio[stPrecio];
			//						break;
			//						#endregion Precio por Cliente
			//						#region Precio por cantidad
			//					case 1: // Precio definido por Cantidad Vendida
			//						double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			//						if (dCantidad >= miArt.Cantidad[5])
			//							fPrecio = (double) miArt.Precio[5];
			//						else if (dCantidad >= miArt.Cantidad[4])
			//							fPrecio = (double) miArt.Precio[4];
			//						else if (dCantidad >= miArt.Cantidad[3])
			//							fPrecio = (double) miArt.Precio[3];
			//						else if (dCantidad >= miArt.Cantidad[2])
			//							fPrecio = (double) miArt.Precio[2];
			//						else //(dCantidad >= (int) miArt.Cantidad[1])
			//							fPrecio = (double) miArt.Precio[1];
			//						break;
			//						#endregion Precio por cantidad
			//						#region Precion por Forma de pago
			//					case 2: // Precio definido por Forma de Pago
			//						C1DataRow drFormaPago = 
			//							cdsCompraTabla.TableViews["CompraFormaPago"].Rows.Find((int) cmbFormaPago.Value);
			//						stPrecio = 1;
			//						if (drFormaPago != null)
			//							stPrecio = (int) drFormaPago["idPrecio"];
			//						fPrecio = (double) miArt.Precio[stPrecio];
			//						break;
			//						#endregion Precion por Forma de pago
			//						#region Precios por Procedimientos
			//					case 3: // Precio por listas
			//						double dCantidadV = (double) e.Cell.Row.Cells["Cantidad"].Value;
			//						int IdClienteV = (int) cmbCliente.Value;
			//						string stExec = "Exec ArticuloPrecioLista " + IdClienteV.ToString() + ", "
			//							+ IdArt + ", " + dCantidadV.ToString("0.00", us);
			//						fPrecio = Funcion.dEscalarSQL(cdsSeteoF, stExec, true);
			//						break;
			//						#endregion Precios por Procedimientos
			//				}
			//					break;
			//				case 4: // Compra
			//				case 2: // Orden de Compra
			//				case 6: // Dev Compra
			//				case 10: // Consumo Real
			//				case 12: // Producto Terminado
			//				case 14: // Reservacion de Proveedor
			//				case 16: // Pedido Proveedor
			//				case 36: // Requisicion
			//					if ((int) cmbTipoCompra.Value == 12)
			//						e.Cell.Row.Cells["Impuesto"].Value = 0;
			//					if ((bool) drSeteoF["c_u_c_p"])
			//						fPrecio = miArt.CostoUltimo;
			//					else
			//						fPrecio = 0;
			//					break;
			//			}
			//			e.Cell.Row.Cells["Precio"].Value = fPrecio;
			#endregion Borrar
		}

		private void DescuentoPorcentaje(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Variables
			if ((double) e.Cell.Row.Cells["DescuentoPorc"].Value < 0.001) return;
			double dDescuentoPorc = (double) e.Cell.Row.Cells["DescuentoPorc"].Value;
			double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			double dIva = (double) e.Cell.Row.Cells["Impuesto"].Value;
			double dPrecio = 0;
			if (e.Cell.Row.Cells["Precio"].Value != DBNull.Value) dPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
			#endregion Variables
			#region Mensajes
			string stExec = string.Format("Exec ArticuloDescPorc {0}, {1}, {2}, {3}, {4}",
				dDescuentoPorc, IdTipoFactura, dCantidad.ToString("0.000000", us), 
				dPrecio.ToString("0.00000", us), dIva.ToString("0.00", us));
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);
			if (stMensaje.Length > 0)
			{
				e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				MessageBox.Show("Descuento Supera el Limite de Descuento Permitido", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			#endregion Mensajes
			#region Valor
			e.Cell.Row.Cells["DescuentoArt"].Value = Funcion.dEscalarSQL(cdsSeteoF, stExec + ", 1");
			#endregion 
		}

		private void Promocion(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return;
			double fCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			int IdArt = (int) e.Cell.Row.Cells["idArticulo"].Value;
			
			string stSelect = "Select TipoPromocion From Articulo Where idArticulo = " + IdArt;
			int ITipoProm = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
			string stExec = string.Format("Exec ArticuloPromocionLista {0}, {1}, {2}, {3} ",
				IdTipoFactura, IdArt, fCantidad.ToString("0.000000", us), (int) rdgContadoCred.Value);
			double dDescuento = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			if (dDescuento <= 0) return;
			if (ITipoProm == 2) // Descuento por Valor
			{
				e.Cell.Row.Cells["DescuentoArt"].Value = dDescuento;
			} 
			if (ITipoProm == 3) // Descuento por Porcentaje
			{
				e.Cell.Row.Cells["DescuentoPorc"].Value = dDescuento;
				double fPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
				e.Cell.Row.Cells["DescuentoArt"].Value = dDescuento * fPrecio * fCantidad / 100.0;
			}
		}


		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Bodega"].Value = (int) spnBodega.Value;
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
				MessageBox.Show(stMensaje); // "No puede crear más lineas");
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
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

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true;
				string stError = e.DataErrorInfo.InvalidValue.ToString();
				string stCol = e.DataErrorInfo.Cell.Column.ToString();
				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol, 
					"Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["CodArt"].Value = e.Row.Cells["IdArticulo"].Value;
			e.Row.Cells["UnidadBotrosa"].Value = e.Row.Cells["IdUnidad"].Value;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Visibilidad de Columnas 
			e.Layout.Bands[0].Columns["Bodega"].Hidden = !(bool) drSeteoF["VerBodega"];
			e.Layout.Bands[0].Columns["CantidadPeso"].Hidden = !(bool) drSeteoF["VerPeso"];
			e.Layout.Bands[0].Columns["CodArt"].Hidden = !(bool) drSeteoF["EscogerCodigo"];
			e.Layout.Bands[0].Columns["DescuentoArt"].Hidden = !(bool) drSeteoF["VerDescuento"];
			e.Layout.Bands[0].Columns["DescuentoPorc"].Hidden = !(bool) drSeteoF["VerDescPorc"];
			e.Layout.Bands[0].Columns["Flete"].Hidden = !(bool) drSeteoF["VerFlete"];
			e.Layout.Bands[0].Columns["idProyecto"].Hidden = !(bool) drSeteoF["VerProyecto"];
			e.Layout.Bands[0].Columns["idSubProyecto"].Hidden = !(bool) drSeteoF["VerSubProyecto"];
			e.Layout.Bands[0].Columns["Notas"].Hidden = !(bool) drSeteoF["VerNotas"];
			e.Layout.Bands[0].Columns["RefCodigo"].Hidden = !(bool) drSeteoF["VerRefCodigo"];
			e.Layout.Bands[0].Columns["RefNumero"].Hidden = !(bool) drSeteoF["VerRefNumero"];
			if (drSeteoF["Importacion"] != DBNull.Value && (bool) drSeteoF["Importacion"])
				e.Layout.Bands[0].Columns["idImportacion"].Hidden = false;

			e.Layout.Bands[0].Columns["Pedido"].Hidden = !(bool) drSeteoF["VerPedido"];
			e.Layout.Bands[0].Columns["Vencimiento"].Hidden = !(bool) drSeteoF["VerVencimiento"];

			#region Lotes
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FPCLOT'"))
			{
				bLote = true;
				e.Layout.Bands[0].Columns["Elaboracion"].Hidden = false;
				e.Layout.Bands[0].Columns["Vencimiento"].Hidden = false;
				e.Layout.Bands[0].Columns["RefCodigo"].Hidden = false;
				e.Layout.Bands[0].Columns["RefCodigo"].Header.Caption = "Lote";
				if (!(IdTipoFactura == (int) Funcion.TipoFactura.Compra || IdTipoFactura == (int) Funcion.TipoFactura.IngresoBodega))
				{
					e.Layout.Bands[0].Columns["RefCodigo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
					e.Layout.Bands[0].Columns["Elaboracion"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
					e.Layout.Bands[0].Columns["Vencimiento"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				}
			}
			#endregion Lotes
			#endregion Visibilidad de Columnas Bodega, CodArt, Ref Codigo, Ref Numero
			#region LimiteTrasnsforma Preguntar Monica
			if (drSeteoF["LimiteTransforma"] != DBNull.Value && (bool) drSeteoF["LimiteTransforma"] == true)
			{
				string stRefC = "Ref. Codigo";
				string stRefN = "Ref. Numero";
				if (IdTipoFactura == 4)
				{
					stRefC = "Orden Compra";
					stRefN = "Requisicion";
				}
				if (IdTipoFactura == 2)
				{
					stRefC = "Compra";
					stRefN = "Requisicion";
				}
				if (IdTipoFactura == 36)
				{
					stRefC = "Compra";
					stRefN = "Orden Compra";
					chkRevisado.Text = "Centralizado";
				}

				e.Layout.Bands[0].Columns["RefCodigo"].Header.Caption = stRefC;
				e.Layout.Bands[0].Columns["RefNumero"].Header.Caption = stRefN;
			}
			#endregion LimiteTrasnsforma Preguntar Monica
			#region Unidades
			e.Layout.Bands[0].Columns["idUnidad"].Hidden = !(bool) drSeteoF["VerUnidad"];
			e.Layout.Bands[0].Columns["Unidades"].Hidden = !(bool) drSeteoF["VerUnidad"];
			if ((bool) drSeteoF["VerUnidad"])
				e.Layout.Bands[0].Columns["Cantidad"].CellActivation = Activation.Disabled;
			#endregion Unidades
			#region Lista desplegable de Articulos
			if (drSeteoF["VerDesplegable"] != DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				e.Layout.Bands[0].Columns["Articulo"].Hidden = false;
				e.Layout.Bands[0].Columns["Codigo"].Hidden = false;
				e.Layout.Bands[0].Columns["CodArt"].Hidden = true;
				e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;
			}
			#endregion Lista desplegable de Articulo
			#region Cambio de Precio
			if (IdTipoFactura == 1 || IdTipoFactura == 3 || IdTipoFactura == 8)
			{
				cmbPrecio.Visible = true;
				btCambiarPrecio.Visible = true;
			}

			if (!(bool) drSeteoF["CambiarPrecio"] && (IdTipoFactura == 1 || IdTipoFactura == 3))
			{
				e.Layout.Bands[0].Columns["Precio"].CellActivation = Activation.Disabled;
				cmbPrecio.Visible = false;
				btCambiarPrecio.Visible = false;
			}
			#endregion Cambio de Precio
			#region Iva

			// Compra Egreso Ingreso
			if (IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 9)
			{
				if (drSeteoF["CambiarIvaC"] == DBNull.Value || !(bool)drSeteoF["CambiarIvaC"])
					e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;
			}	
			else if (!(bool)drSeteoF["CambiarIva"])
				e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;

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
			this.txtIceTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtIva0.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			this.txtDescIva.FormatString = stFormato;
			this.txtDescIva0.FormatString = stFormato;
			this.txtDescIva0Total.FormatString = stFormato;
			this.txtDescLinea.FormatString = stFormato;
			this.txtDescLinea0.FormatString = stFormato;
			this.txtFlete.FormatString = stFormato;

			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIceTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtIva0.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			this.txtDescIva.MaskInput = stInput;
			this.txtDescIva0.MaskInput = stInput;
			this.txtDescIva0Total.MaskInput = stInput;
			this.txtDescLinea.MaskInput = stInput;
			this.txtDescLinea0.MaskInput = stInput;
			this.txtFlete.MaskInput = stInput;
			#endregion
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].Format = stFormato;
			e.Layout.Bands[0].Columns["CantidadPeso"].Format = stFormato;
			e.Layout.Bands[0].Columns["Pedido"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["CantidadPeso"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Pedido"].MaskInput = stInput;
			#endregion
			#region Decimales para el precio
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;
			#endregion
			#region Decimales para el Impuesto
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;
			#endregion
			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPorcDesc"]);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			this.txtDescPorcIva.FormatString = stFormato;
			this.txtDescPorcIva0.FormatString = stFormato;
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, iDecPorD);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			this.txtDescPorcIva.MaskInput = stInput;
			this.txtDescPorcIva0.MaskInput = stInput;
			#endregion
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
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
						if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "REFCODIGO" && barraDato1.bEditar)
						{
							int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
							double Cantidad = (double) ultraGrid1.ActiveRow.Cells["Cantidad"].Value;
							
							using (Lotes miLote = new Lotes(IdArt, (int)Cantidad))
							{
								if (miLote.ShowDialog() == DialogResult.OK)
								{
									if (!(IdTipoFactura == (int) Funcion.TipoFactura.Compra
										|| IdTipoFactura == (int) Funcion.TipoFactura.IngresoBodega))
									{
										ultraGrid1.ActiveRow.Cells["RefCodigo"].Value = miLote.stLote;
										ultraGrid1.ActiveRow.Cells["Vencimiento"].Value = miLote.dtVencimiento;
										ultraGrid1.ActiveRow.Cells["Elaboracion"].Value = miLote.dtElaboracion;
									}
								}
							}
						}
						if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "REFNUMERO")
						{
							int IdDetCompra = (int) ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
							if (barraDato1.bNuevo || IdDetCompra < 0)
							{
								MessageBox.Show("Antes de ingresar la serie debe grabar el registro",
									"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
							if (IdArt < 1)
							{
								MessageBox.Show("Ingrese Articulo",
									"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							int IdCompra = (int) txtIdCompra.Value;
							double dCantidad = (double) ultraGrid1.ActiveRow.Cells["Cantidad"].Value;
							int iCantidad = int.Parse(dCantidad.ToString("0", us));
							int iDisponible = 0;
							// Compra o Ingreso de Bodega
							if (IdTipoFactura == 4 || IdTipoFactura == 9)
							{
								if (!e.Shift)
								{
									string stExec = "Exec SerieCrear " + IdArt.ToString() + ", "
										+ IdCompra.ToString() + ", " + iCantidad.ToString()
										+ ", 1, " + IdDetCompra.ToString();

									Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
								}
							}
							// Venta o Egreso de Bodega
							if (IdTipoFactura == 1 || IdTipoFactura == 8)
							{
								if (!e.Shift)	iDisponible = iCantidad;
							}
							// Devolucion de Venta
							int idTipoFactura = 0;
							if (IdTipoFactura == 5)
							{
								if (!e.Shift)	iDisponible = -iCantidad;
								idTipoFactura = 1;
							}
							// Devolucion de Compra
							if (IdTipoFactura == 6)
							{
								if (!e.Shift)	iDisponible = iCantidad;
								idTipoFactura = 4;
							}
							bool bUsado = e.Shift;
							CompraSerie miSerie = new CompraSerie(IdArt, IdCompra, iDisponible, 
								idTipoFactura, IdDetCompra, bUsado);
							miSerie.ShowDialog();
						}
						if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "PRECIO")
						{
							int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
							if (IdArt < 1)
							{
								MessageBox.Show("Ingrese Articulo",
									"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							CompraArt miCompraArt = new CompraArt(IdArt);
							miCompraArt.ShowDialog();
						}
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

		#endregion Ultragrid
		#region Botones principales
		#region Boton Cliente
		private void btCliente_Click(object sender, System.EventArgs e)
		{
			bool bCliente = !Funcion.EsProveedor(IdTipoFactura);
			int IdCliente = 0;
			Cliente miCliente;
			if (cmbCliente.Value == DBNull.Value)
				miCliente = new Cliente(bCliente);
			else
			{
				IdCliente = (int) cmbCliente.Value;
				miCliente = new Cliente(bCliente, IdCliente);
			}

			miCliente.ShowDialog();
			if (drSeteoF["VerNombreLista"] != DBNull.Value && (bool) drSeteoF["VerNombreLista"])
			{
				DateTime dtIni = DateTime.Now;
				string stExec = "Exec ClienteCargaNombre " + IdTipoFactura;
				cmbCliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				cmbCliente.Refresh();
				TimeSpan ts = DateTime.Now.Subtract(dtIni);
				lblTiempo.Text = "T: " + ts.TotalSeconds.ToString("0.0000", us);
			}
			if (barraDato1.bNuevo && IdCliente == 0)
			{
				try
				{
					cmbCliente.Value = int.Parse(miCliente.txtIdCliente.Value.ToString());
				}
				catch{};
			}
			miCliente.Dispose();
			miCliente = null;
			cmbCliente.Select();
		}

		private void btCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			#region Saldo de Cliente
			string stTipoSaldo = "";
			string stExec = string.Format("Exec ClienteSaldo {0}, 1", drCompra["idCliente"]);
			double dSaldoCxC = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			if (dSaldoCxC >= 0.01) stTipoSaldo += "Saldo por cobrar: " + dSaldoCxC.ToString("#,##0.00", us) + "\n";
			stExec = string.Format("Exec ClienteSaldo {0}, 4", drCompra["idCliente"]);
			dSaldoCxC = Funcion.dEscalarSQL(cdsSeteoF, stExec);
			if (dSaldoCxC >= 0.01) stTipoSaldo += "Saldo por pagar: " + dSaldoCxC.ToString("#,##0.00", us) + "\n";
			#endregion Saldo de Cliente
			#region Variacion de saldo de factura
			if (!barraDato1.bNuevo)
			{
				double dSaldoCompra = (double)drCompra["Saldo"];
				double dTotalCompra = (double)drCompra["Total"];
				if (Math.Abs(dTotalCompra - dSaldoCompra) >= 0.01)
				{
					dTotalCompra -= dSaldoCompra;
					stTipoSaldo += "Variacion de Saldo: " + dTotalCompra.ToString("#,##0.00", us);
				}
			}
			#endregion Variacion de saldo de factura
			#region Vencimiento
			double dVencido = SaldoMora();
			if (dVencido > 0) 
			{
				stTipoSaldo += "\nVencido: " + dVencido.ToString("#,##0.00", us);
			}
			#endregion Vencimiento
			if (stTipoSaldo.Length == 0) stTipoSaldo = "No existen Saldos";
			MessageBox.Show(stTipoSaldo, "Informacion de Saldos",
				MessageBoxButtons.OK, MessageBoxIcon.Information);	
		}
		private double SaldoMora()
		{
			string stSelect = "Exec ClienteSaldoVencido " + IdTipoFactura;
			double dSaldo = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
			return dSaldo;
		}

		#endregion Boton Cliente
		#region Cliente Modificado
		private void cmbCliente_Validated(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				miProveedor = new Proveedor(cdsSeteoF, (int) drCompra["idCliente"]);
			#region Nuevo Registro
			if (barraDato1.bNuevo) 
			{
				Autorizaciones();
				if ((bool) drSeteoF["Copia_Dist"]) ClientePredefinidos();
			}
			#endregion Nuevo Registro
		}
		private void Autorizaciones()
		{
			if ((IdTipoFactura == 4 || IdTipoFactura == 2) && drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
			{
				drCompra["AutFactura"] = miProveedor.AutorizaFactura;
				drCompra["SerieFactura"] = miProveedor.SerieFactura;
				if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-60))
					drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
				if (miProveedor.idCredTribut > 0)
					drCompra["idCredTributario"] = miProveedor.idCredTribut;
			}
		}

		private void ClientePredefinidos()
		{
			if ((int) rdgContadoCred.Value == 2)
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
			}
			else
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
			}

			drCompra["idVendedor"] = miProveedor.idVendedor;
			drCompra["DiasCredito"] = miProveedor.DiasCredito;
		}

		private void cmbCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stPrimero = "Nombre";
				if (cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "Codigo";
				cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbCliente.DisplayMember = stPrimero;
			}
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stFiltro = "0";
			if (Funcion.EsProveedor(IdTipoFactura))
				stFiltro = "1";

			string stNombre = txtNombre.Text.Trim();
			if (stNombre.Length == 0) return;
			string stSelect = "Select Count(*) From Cliente Where Nombre like '"
				+ stNombre + "%' And (Ambos = 1 Or Proveedor = " + stFiltro + ")";
			int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
			int IdCliente = 0;
			if (iCuenta != 1)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(stNombre, IdTipoFactura))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						IdCliente = (int) miBusqueda.txtIdBusca.Value;
						drCompra["Nombre"] = miBusqueda.txtDescripcion.Text;
					}
					else e.Cancel=true;
				}
			}
			else
			{
				stSelect = "Select idCliente From Cliente Where Nombre like '"
					+ stNombre + "%' And (Ambos = 1 Or Proveedor = " + stFiltro + ")";
				IdCliente = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
				stSelect = "Select Nombre From Cliente Where idCliente = " + IdCliente.ToString();
				drCompra["Nombre"] = Funcion.sEscalarSQL(cdsSeteoF, stSelect, true);
			}
			drCompra["idCliente"] = IdCliente;

			if (!barraDato1.bEditar) return;
			if (drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				miProveedor = new Proveedor(cdsSeteoF, (int) drCompra["idCliente"]);
			#region Nuevo Registro
			if (barraDato1.bNuevo) 
			{
				if ((IdTipoFactura == 4 || IdTipoFactura == 2) && drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				{
					drCompra["AutImprenta"] = miProveedor.AutorizaImprenta;
					drCompra["SerieFactura"] = miProveedor.SerieFactura;
					drCompra["AutFactura"] = miProveedor.AutorizaFactura;
					if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-10))
						drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
					if (miProveedor.idCredTribut > 0)
						drCompra["idCredTributario"] = miProveedor.idCredTribut;
				}
				if ((bool) drSeteoF["Copia_Dist"] == false) return;
				if ((int) rdgContadoCred.Value == 2)
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
				}
				else
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
				}

				drCompra["idVendedor"] = miProveedor.idVendedor;
				drCompra["DiasCredito"] = miProveedor.DiasCredito;
			}
			#endregion Nuevo Registro
		}

		#endregion Cliente
		#region Boton Comprobante
		private void cmbComprobante1_Validated(object sender, System.EventArgs e)
		{
			// Nota de Venta o Documento emitido en el exterior no sustenta credito tributario
			if ((int) cmbComprobante1.Value == 2 || (int) cmbComprobante1.Value == 12 || (int) cmbComprobante1.Value == 18) //Reembolso Gastos
				drCompra["idCredTributario"] = 2;
			else
				drCompra["idCredTributario"] = 1;
		}

		private void cmbComprobante1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (CompraComprobante miComprob = new CompraComprobante())
			{
				miComprob.ShowDialog();
			}
			cdsCompraTabla.Clear();
			cdsCompraTabla.Fill();
			cmbComprobante1.Refresh();
			cmbComprobante1.Select();
		}

		#endregion Boton Comprobante
		#region Boton Total
		private void cmdTotal()
		{
			#region Creacion de Variables
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
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				dIva = 0;
				dDescuento = 0;
				if (dr.Cells["Cantidad"].Value != DBNull.Value) dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != DBNull.Value)	dPrecio = (double) dr.Cells["Precio"].Value;
				if (dr.Cells["Impuesto"].Value != DBNull.Value)	dIva = (double) dr.Cells["Impuesto"].Value;
				if (dr.Cells["DescuentoArt"].Value != DBNull.Value)	dDescuento = (double) dr.Cells["DescuentoArt"].Value;
				dSubTotal = dCantidad * dPrecio + .000001;
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
			if (txtDescPorcIva.Value != DBNull.Value && (double) txtDescPorcIva.Value > 0)
				drCompra["Desc1"] = (double) txtDescPorcIva.Value * (dSubtIva - dDescIva) / 100;
			if (txtDescPorcIva0.Value != DBNull.Value && (double) txtDescPorcIva0.Value > 0)
				drCompra["Desc2"] = (double) txtDescPorcIva0.Value * (dSubtIva0 - dDescIva0) / 100;
			drCompra["Desc3"] = dDescIva;
			drCompra["Desc4"] = dDescIva0;

			dDescuento = dDescIva0 + dDescIva;
			dDescIva = (double) drCompra["Desc1"];
			dDescIva0 = (double) drCompra["Desc2"];
			drCompra["Descuento"] = dDescIva + (double) drCompra["Desc3"];
			drCompra["Descuento0"] = dDescIva0 + (double) drCompra["Desc4"];
			dDescuento += dDescIva0 + dDescIva; // descuento total de la factura
			dDescuento = Math.Round(dDescuento, dRedondeo);
			#endregion Calculo de Descuentos

			#region Calculo de subtotales
			dSubtIva0 += (double) txtFlete.Value;
			dSubtIva0 = Math.Round(dSubtIva0, dRedondeo);
			dSubtIva = Math.Round(dSubtIva, dRedondeo);
		
			drCompra["SubTotalExcento"] = dSubtIva0;
			drCompra["SubTotalIva"] = dSubtIva;
			//			if ((bool) drSeteoF["IncluyeIva"])
			//				drCompra["SubTotalIva"] = dSubtIva;
			//			else
			//				drCompra["SubTotalIva"] = dSubtIva;

			#region no se que significa revisar posteriormente
			double dSubtIvaAntes = dSubtIva - (double) drCompra["Desc3"];
			if ((double) drCompra["Desc1"] > 0 && dSubtIvaAntes > 0)
				dImpuesto -= dImpuesto * (1 - (dSubtIvaAntes - (double) drCompra["Desc1"])/ dSubtIvaAntes);
			dImpuesto = Math.Round(dImpuesto, dRedondeo);
			#endregion no se que significa revisar posteriormente
			drCompra["Iva"] = dImpuesto;
			#region Servicio
			if (txtServicio.Width > 0 && IdTipoFactura == 1 && bServicioAuto) 
			{
				txtServicio.Value = (dSubtIva0 + dSubtIva- dDescuento) * 0.1 ;
				drCompra["Servicio"] = (dSubtIva0 + dSubtIva- dDescuento) * 0.1 ;
			}
			#endregion Servicio
			#endregion Calculo de subtotales

			#region Calculo del Total
			double dTotal = dSubtIva0 + dSubtIva + dImpuesto - dDescuento 
				+ (double) txtIceTotal.Value 
				+ (double) txtServicio.Value
				+ (double) txtPropina.Value;

			dTotal = Math.Round(dTotal, dRedondeo);
			if(dTotal <0)	
			{
				drCompra["Total"] = 0;
				errorProvider.SetError(txtTotal, "Error en el calculo del total: " 
					+ dTotal.ToString("#,##0.00"));
			}
			else
			{
				drCompra["Total"] = dTotal;
				errorProvider.SetError(txtTotal, ""); 
			}
			#endregion Calculo del Total
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			cmdTotal();
		}

		#endregion Boton Total

		#endregion Fin de Botones Principales
		#region Pestanas
		#region Notas
		#region Centro de Costo
		private void btCentro_Click(object sender, System.EventArgs e)
		{
			using (Proyecto miProyecto = new Proyecto())
			{
				miProyecto.ShowDialog();
			}
			string stExec = "Exec ProyectoCarga " + IdTipoFactura;
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbProyecto1.DataSource = cmbProyecto.DataSource;
		}
		private void btCentro_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Tiene que estar grabada la factura para repartir centros de costo",
					"Información");
				return;
			}
			if (drCompra["idProyecto"] == DBNull.Value || (int) drCompra["idProyecto"] == 0)
			{
				MessageBox.Show("No puede pasar Centros de Costo vacios a las lineas", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea pasar los Centros de Costo a las líneas?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int IdProyecto = (int) drCompra["idProyecto"];
			string stExec = "Exec CentroCostoCompra " + txtIdCompra.Value;
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
			barraDato1.ProximoId(5);
		}

		private void cmbProyecto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRPBP'")) 				
			{				
				if ((int) cmbProyecto.Value > 0)
					cmbSubProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
						"Exec SubProyectoFiltro " + cmbProyecto.Value);
			}
		}

		private void cmbProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (ultraGrid1.ActiveRow != null)
					ultraGrid1.ActiveRow.Cells["idProyecto"].Value = 0;
				drCompra["idProyecto"] = 0;
			}
		}

		#endregion Centro de Costo
		#region Proyecto
		private void btSubProyecto_Click(object sender, System.EventArgs e)
		{
			using (SubProyecto miProyecto = new SubProyecto())
			{
				miProyecto.ShowDialog();
			}
			string stExec = "Exec SubProyectoCarga " + IdTipoFactura;
			cmbSubProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			cmbSubProyecto1.DataSource = cmbSubProyecto.DataSource;
		}

		private void btSubProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Tiene que estar grabada la factura para repartir Proyectos",
					"Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea pasar los Proyectos a las líneas?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			if (drCompra["idSubProyecto"] == DBNull.Value || (int) drCompra["idSubProyecto"] == 0)
			{
				MessageBox.Show("No puede pasar Proyectos vacios a las lineas", "Información");
				return;
			}
			string stExec = "Exec ProyectoCompra " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteoF, stExec, true);
			barraDato1.ProximoId(5);
		}

		private void cmbSubProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (ultraGrid1.ActiveRow != null)
					ultraGrid1.ActiveRow.Cells["idSubProyecto"].Value = 0;

				cmbSubProyecto.Value = 0;
				cmbSubProyecto.Select();
				cmbSubProyecto.Refresh();
			}
		}

		#endregion SubProyecto
		#region Comprobante
		private void btComprobante_Click(object sender, System.EventArgs e)
		{
			if (txtComprobante.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Comprobante", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			CompraComprobanteBusca(txtComprobante.Text.Trim());
		}

		private void CompraComprobanteBusca(string stComprobante)
		{
			int idCliente = (int) cmbCliente.Value;
			string stExec = string.Format("Exec CompraClienteComprobante {0}, {1}, '{2}'",
				idCliente, IdTipoFactura, stComprobante);
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);
			if (stMensaje.StartsWith("-"))
			{
				int IdCompra = int.Parse(stMensaje);
				Compra miCompra = new Compra(IdCompra);
				miCompra.MdiParent = MdiParent;
				miCompra.Show();
			}
			else
			{
				MessageBox.Show(stMensaje, "Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}	
		}

		private void btComprobante_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if ((e.Button == MouseButtons.Right) && 
				(IdTipoFactura == 4 || IdTipoFactura == 9 || IdTipoFactura == 14 || IdTipoFactura == 31))
			{
				Cursor = Cursors.WaitCursor;
				string stExec = "Exec OrdenCompraConsumo '";
				if (drSeteoF["OrdenC_ConsumoL"] != DBNull.Value && (bool) drSeteoF["OrdenC_ConsumoL"] == true)
					stExec = "Exec OrdenCompraConsumoPedProv '";
				stExec += txtComprobante.Text.Trim() + "'";
				Funcion.EjecutaSQL(cdsCompra, stExec, true);
				new Reporte("OrdenCompraConsumo.rpt").Show();
				Cursor = Cursors.Default;
			}
		}

		#endregion Comprobante
		#region Vendedor
		private void cmbVendedor_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbVendedor.Value = 0;
			cmbVendedor.Select();
			cmbVendedor.Refresh();
		}

		private void btVendedor_Click(object sender, System.EventArgs e)
		{
			int idPersonal = 0;
			if (drCompra["idVendedor"] != DBNull.Value) idPersonal = (int) drCompra["idVendedor"];
			RolPersonal miPersonal = new RolPersonal(idPersonal);
			miPersonal.ShowDialog();
			string stExec = "Exec PersonalVendedor " + IdTipoFactura;
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
		}

		#endregion Vendedor
		#endregion Notas
		#region Pago
		private void btPagos_Click(object sender, System.EventArgs e)
		{
			// Pagos solo existen en ciertos tipos de comprobantes
			if (IdTipoFactura != (int) Funcion.TipoFactura.Venta && IdTipoFactura != (int) Funcion.TipoFactura.Compra &&
				IdTipoFactura != (int) Funcion.TipoFactura.DebitoVenta && IdTipoFactura != (int) Funcion.TipoFactura.DebitoCompra &&
				IdTipoFactura != (int) Funcion.TipoFactura.DevolucionVenta && IdTipoFactura != (int) Funcion.TipoFactura.DevolucionCompra &&
				IdTipoFactura != (int) Funcion.TipoFactura.ReservaCliente && IdTipoFactura != (int) Funcion.TipoFactura.ReservaProveedor 
				&& IdTipoFactura != 99) return;

			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede crear pagos mientras edita la factura",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((int) rdgContadoCred.Value != 2) 
			{
				MessageBox.Show("Pagos solo se pueden realizar en Créditos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Pagos miPago = new Pagos(IdTipoFactura, (int) txtIdCompra.Value);
			miPago.MdiParent = MdiParent;
			miPago.Show();
		}

		private void btPago_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Para generar un pago debe grabar previamente el registro", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			string stExec = "Exec AbonoCruzeVerifica " + txtIdCompra.Value;
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea Generar un abono?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			stExec = "Exec AbonoCruzePago " + txtIdCompra.Value;
			stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec, true);

			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Informacion de Error",
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			MessageBox.Show("Abono Generado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);	
		}

		private void btFormaPago_Click(object sender, System.EventArgs e)
		{
			CompraFormaPago miFormaPago = new CompraFormaPago();
			miFormaPago.ShowDialog();
			cdsCompraTabla.Clear();
			cdsCompraTabla.Fill();
		}

		private void cmbFormaPago_Validated(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (IdTipoFactura == (int) Funcion.TipoFactura.Venta || IdTipoFactura == (int) Funcion.TipoFactura.Cotizacion)
			{
				if (cmbFormaPago.ActiveRow.Cells["idFormaPago"].Value == DBNull.Value) return;
				int idFormaPago = (int) cmbFormaPago.ActiveRow.Cells["idFormaPago"].Value;
				string stSelect = "Select IsNull(Descuento, -1) From CompraFormaPago Where idFormaPago = " + idFormaPago;
				double dDescuento = Funcion.dEscalarSQL(cdsCompra, stSelect);
				if (dDescuento < 0) return;
				if (dDescuento == 0) drCompra["Desc1"] = 0;
				drCompra["DescPorcIva"] = dDescuento;
				btRefrescaI_Click(this, e);
				cmdTotal();
			}
		}

		private void btCambiarPrecio_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (!miAcceso.CambiarPrecio) return;
			if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value == DBNull.Value
				|| (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value < 1) return;
			if (cmbPrecio.Value == null) 
			{
				MessageBox.Show("Seleccione la lista de precios para cambiar en el artículo seleccionado",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

			string stPrecio = "Precio" + cmbPrecio.Value.ToString().Trim();
			string stSelectPrecio = "Select " + stPrecio + " From Articulo Where idArticulo = " + IdArt;
			ultraGrid1.ActiveRow.Cells["Precio"].Value = Funcion.dEscalarSQL(cdsSeteoF, stSelectPrecio, true);
			ultraGrid1.ActiveRow.Cells["Precio"].Activate();
			ultraGrid1.ActiveRow.Cells["Precio"].Refresh();
			ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
			ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);

			cmdTotal();
			MessageBox.Show("Precio Modificado", "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btCambiarPrecio_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!barraDato1.bEditar) return;

			#region No tiene acceso a  cambiar precio
			if (!miAcceso.CambiarPrecio)
			{
				int IdArt = int.Parse(ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString());
				string stExec = "Exec ArticuloCambiarPrecio1 " + IdArt;
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);

				if (stMensaje.StartsWith("No")) 
				{
					MessageBox.Show(stMensaje, "Información", 
						MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				else
				{
					#region Llamada a pantalla de cambio de precio
					using (CambiarPrecio miCambio = new CambiarPrecio(stMensaje))
					{
						if (DialogResult.OK == miCambio.ShowDialog())
						{
							ultraGrid1.ActiveRow.Cells["Precio"].Value = miCambio.txtPrecio.Value;
							MessageBox.Show("Cambio de precio Aceptado",
								"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
					#endregion Llamada a pantalla de cambio de precio
				}
				return;
			}
			#endregion No tiene acceso a  cambiar precio

			#region Cambia precio de toda la lista
			#region Verifica si ha escogido un precio de la grilla
			if (cmbPrecio.Value == null) 
			{
				MessageBox.Show("Ingrese Precio al que desea Actualizar la lista",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				cmbPrecio.Select();
				return;
			}
			#endregion Verifica si ha escogido un precio de la grilla
			#region Recorre la grilla cambiando los precios
			string stPrecio = "Precio" + cmbPrecio.Value.ToString().Trim();
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == DBNull.Value) continue;
				int IdArt = (int) dr.Cells["idArticulo"].Value;
				if (IdArt < 1) return;
				string stSelectPrecio = "Select " + stPrecio + " From Articulo Where idArticulo = " + IdArt;
				dr.Cells["Precio"].Value = Funcion.dEscalarSQL(cdsSeteoF, stSelectPrecio, true);
				dr.Cells["Precio"].Activate();
			}
			#endregion Recorre la grilla cambiando los precios
			#region Recalculo de Total y mensaje
			cmdTotal();
			MessageBox.Show("Precios Modificados", "Información", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
			#endregion Recalculo de Total y mensaje
			#endregion Cambia precio de toda la lista
		}

		private void cmdVuelto_Click(object sender, System.EventArgs e)
		{
			float fTotal, fVuelto, fRecibido;
			if (spnVuelto.Value == DBNull.Value) return;
			fRecibido = float.Parse(spnVuelto.Value.ToString());
			fTotal = float.Parse(txtTotal.Value.ToString());
			if (fRecibido < fTotal)
			{
				MessageBox.Show("Valor recibido no puede ser menor que total a cobrar",
					"Error en Valor Recibido", MessageBoxButtons.OK, 
					MessageBoxIcon.Error);
				return;
			}
			fVuelto = fRecibido - fTotal;
			MessageBox.Show("Su Cambio es de: " + fVuelto.ToString("#,##0.00"),
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void rdgContadoCred_Validated(object sender, System.EventArgs e)
		{
			if (!barraDato1.bNuevo) return;
			if ((bool) drSeteoF["Copia_Dist"] == false) return;
			if ((int) rdgContadoCred.Value == 2)
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
			}
			else
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
			}
		}

		#endregion Pago
		#region Descuento
		private void btDescuento_Click(object sender, System.EventArgs e)
		{
			if (IdTipoFactura != 4) return; // Aplica solo a compras
			if (txtDescIva.Value == DBNull.Value || (double) txtDescIva.Value == 0) return;
			double dDescuento = double.Parse(txtDescIva.Value.ToString());
			double dDescuento0 = 0;
			if (txtDescIva0.Value == DBNull.Value || (double) txtDescIva0.Value == 0)
				dDescuento0 = double.Parse(txtDescIva0.Value.ToString());

			double dSubTotalIva = 0;
			double dSubTotalIva0 = 0;
			if (txtIva.Value != DBNull.Value) dSubTotalIva = double.Parse(txtIva.Value.ToString());
			if (txtIva0.Value != DBNull.Value) dSubTotalIva0 = double.Parse(txtIva0.Value.ToString());
			if (dSubTotalIva + dSubTotalIva0 == 0) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Distribuir Descuento proporcionalmente?", 
				"Confirmación",	MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			double dPorc = dDescuento / (dSubTotalIva);
			double dPorc0 = dDescuento0 / (dSubTotalIva0);
			double dCantidad, dPrecio, dImpuesto;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				dImpuesto = 0;
				if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
					dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != System.DBNull.Value)
					dPrecio = (double) dr.Cells["Precio"].Value;
				if (dr.Cells["Impuesto"].Value != System.DBNull.Value)
					dImpuesto = (double) dr.Cells["Impuesto"].Value;
				if (dImpuesto > 0)
					dr.Cells["DescuentoArt"].Value = dCantidad * dPrecio * dPorc;
				else
					dr.Cells["DescuentoArt"].Value = dCantidad * dPrecio * dPorc0;
			}
			drCompra["Desc1"] = 0;
			drCompra["Desc2"] = 0;
			drCompra["DescPorcIva"] = 0;
			drCompra["DescPorcIva0"] = 0;
			cmdTotal();
			MessageBox.Show("Descuento Repartido", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btFlete_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea calcular el total del flete?", 
				"Confirmación",	MessageBoxButtons.YesNo, 
				MessageBoxIcon.Information)) return;
			double dFlete = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Flete"].Value != DBNull.Value) dFlete += (double) dr.Cells["Flete"].Value;
			}
			drCompra["Flete"] = dFlete;
			MessageBox.Show("Flete Calculado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btFlete_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!barraDato1.bEditar) return;
			//			if (!(IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 3 || IdTipoFactura == 9)) 
			//				return; // Aplica solo a compras, Egreso, Cotizacion
			if (txtFlete.Value == DBNull.Value) return;
			if ((double) txtFlete.Value == 0) return;
			cmdTotal();
			double dFlete = double.Parse(txtFlete.Value.ToString());
			double dSubTotalIva = 0;
			double dSubTotalIva0 = 0;
			if (txtIva.Value != DBNull.Value) dSubTotalIva = double.Parse(txtIva.Value.ToString());
			if (txtIva0.Value != DBNull.Value) dSubTotalIva0 = double.Parse(txtIva0.Value.ToString()) - dFlete;
			if (dSubTotalIva + dSubTotalIva0 == 0) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Distribuir flete proporcionalmente?", 
				"Confirmación",	MessageBoxButtons.YesNo, MessageBoxIcon.Information)) return;
			double dPorc = dFlete / (dSubTotalIva + dSubTotalIva0);
			double dCantidad, dPrecio;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				if (dr.Cells["Cantidad"].Value != DBNull.Value)	dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != DBNull.Value)	dPrecio = (double) dr.Cells["Precio"].Value;
				dr.Cells["Flete"].Activate();
				dr.Cells["Flete"].Value = dCantidad * dPrecio * dPorc;
			}
			MessageBox.Show("Flete Repartido", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btComision_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar comisiones mientras edita", "Información");
				return;
			}
			CompraComision miComision = new CompraComision((int) txtIdCompra.Value);
			miComision.MdiParent = MdiParent;
			miComision.Show();
		}

		private void btCorresponsal_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar corresponsales mientras edita", "Información");
				return;
			}
			CompraCorresponsal miCorresponsal = new CompraCorresponsal((int) txtIdCompra.Value);
			miCorresponsal.MdiParent=this.MdiParent;
			miCorresponsal.Show();
		}

		private void btRefrescaD_Click(object sender, System.EventArgs e)
		{
			double dTotalIva = (double) txtIva.Value; // Valor que paga Iva de la factura
			dTotalIva -= (double) txtDescLinea.Value;
			double dDescIva = (double) txtDescIva.Value;
			double dPorcIva = 0;
			if (dTotalIva != 0) dPorcIva = dDescIva / (dTotalIva) * 100.0;
			this.drCompra["DescPorcIva"] = dPorcIva;

			double dTotalIva0 = (double) txtIva0.Value; // Valor que paga Iva 0 de la factura
			double dDescIva0 = (double) txtDescIva0.Value;
			double dDescIvaLinea0 = (double) txtDescLinea0.Value;
			double dPorcIva0 = 0;
			if (dTotalIva0 != 0) dPorcIva0 = dDescIva0 / (dTotalIva0 - dDescIvaLinea0) * 100.0;
			drCompra["DescPorcIva0"] = dPorcIva0;
		}

		private void btRefrescaI_Click(object sender, System.EventArgs e)
		{
			double dTotalIva = (double) txtIva.Value; // Valor que paga Iva de la factura
			dTotalIva -= (double) txtDescLinea.Value;
			double dDescPIva = (double) txtDescPorcIva.Value;
			double dDescIva = dDescPIva *  dTotalIva / 100.0;
			drCompra["Desc1"] = dDescIva;

			double dTotalIva0 = (double) txtIva0.Value; // Valor que no paga Iva de la factura
			dTotalIva0 -= (double) txtDescLinea0.Value;
			double dDescPIva0 = (double) txtDescPorcIva0.Value;
			double dDescIva0 = dDescPIva0 * dTotalIva0 / 100.0;
			drCompra["Desc2"] = dDescIva0;
		}

		#endregion Descuento
		#region Entrega
		private void btImportacion_Click(object sender, System.EventArgs e)
		{
			string stDepartamento = txtDepartamento.Text.ToString().Trim();
			if (stDepartamento.Length == 0)
			{
				MessageBox.Show("Ingrese Importacion", "Información");
				return;
			}

			Cursor = Cursors.WaitCursor;
			string stFiltro = string.Format("{{Compra.Departamento}} = '{0}' And {{Compra.idTipoFactura}} = {1}",
				stDepartamento, IdTipoFactura);
			if (IdTipoFactura == 14)
				new Reporte("ImportacionArancel", stFiltro).Show();
			else
				new Reporte("NotaEntregaImp", stFiltro).Show();
			Cursor = Cursors.Default;
		}

		private void btImportacion_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (MouseButtons.Right != e.Button) return;
			Cursor = Cursors.WaitCursor;
			string stIdCompra = "{Compra.idCompra} = " + txtIdCompra.Value;
			new Reporte("ImportacionArancel.Rpt", stIdCompra).Show();
			Cursor = Cursors.Default;
		}

		private void btComprob1_Click(object sender, System.EventArgs e)
		{
			if (txtComprob1.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Comprobante", "Información");
			}
			CompraComprobanteBusca(txtComprob1.Text.Trim());
		}

		private void cmbRecibe_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbRecibe.Value = 0;
			cmbRecibe.Select();
			cmbRecibe.Refresh();
		}

		#endregion Entrega
		#region Cuotas
		private void btTotalCuotas_Click(object sender, System.EventArgs e)
		{
			#region Verifica Cuotas
			errorProvider.SetError(txtNumCuotas, "");
			if (txtNumCuotas.Value.ToString().Length == 0)
			{
				errorProvider.SetError(txtNumCuotas, "Ingrese Numero de Cuotas");
				return;
			}
			int iPlazo = int.Parse(txtNumCuotas.Value.ToString());
			if (iPlazo == 0)
			{
				txtValorCuotas.Value = 0;
				txtTotalCuotas.Value = 0;
				txtInteresPagar.Value = 0;
				return;
			}
			#endregion Verifica Cuotas
			#region Calculo de Cuota
			double dMonto = (double) txtTotal.Value - (double) txtAnticipo.Value - (double) drCompra["Retenido"];
			double dInteres = (double) txtPorcInteres.Value / 12 / 100;
			double dCuota = 0;
			if ((bool) drSeteoF["InteresLineal"] == true)
				dCuota = (dMonto * (1 + dInteres * iPlazo))/iPlazo;
			else
			{
				if (dInteres == 0)
					dCuota = dMonto / iPlazo;
				else
					dCuota = dMonto * dInteres / (1-Math.Pow(1 + dInteres, -iPlazo));
			}

			//Caso de interes no calculado
			if ((double) txtPorcInteres.Value == 0 && iPlazo == 1)
				dCuota = dMonto + (double) txtInteresPagar.Value;
			#endregion Calculo de Cuota
			#region Calculo de Totales
			drCompra["ValorCuotas"] = Math.Round(dCuota, 2);
			drCompra["TotalCuotas"] = Math.Round(dCuota, 2) * iPlazo + (double) txtAnticipo.Value;
			if ((double) txtPorcInteres.Value ==0)
			{
				drCompra["TotalCuotas"] = (double) txtTotal.Value;
				drCompra["ValorInteres"] = 0;
			}
			else
			{
				if ((double) txtTotalCuotas.Value < (double) txtTotal.Value)
					drCompra["TotalCuotas"] = (double) txtTotal.Value;
				drCompra["ValorInteres"] = (double) txtTotalCuotas.Value - (double) txtTotal.Value;
			}
			#endregion Calculo de Totales
		}

		private void btTotalCuotas_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			#region Verifica
			if (e.Button != MouseButtons.Right) return;
			if (IdTipoFactura != 1 && IdTipoFactura != 4 && IdTipoFactura != 11) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Debe grabar la factura antes de generar pagos", "Información");
				return;
			}
			if (drSeteoF["PagoInteres"] == DBNull.Value && (bool) drSeteoF["PagoInteres"] == false)
			{
				MessageBox.Show("Para habilitar la opcion de crear tabla de amortizacion active la opcion 'Pago con Interes'"
					+ "\nPantalla Archivo - Propiedades - Facturacion en la pestaña Contabilidad", "Información");
				return;
			}
			#endregion Verifica
			#region Generacion de Tabla de Amortizacion
			string stExec = "Exec CompraAmortizacion " + txtIdCompra.Value;
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stExec);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				return;
			}
			MessageBox.Show("Tabla de Amortizacion Generada", "Información");
			#endregion Generacion de Tabla de Amortizacion
		}

		#endregion Cuotas
		#region A.T.
		private void cmbCredTribut_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using(CompraCredTribut miCred = new CompraCredTribut())
			{
				miCred.ShowDialog();
			}
		}

		private void cmbIce_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (CompraPorcIce miPorcIce = new CompraPorcIce())
			{
				miPorcIce.ShowDialog();
			}
		}
		#endregion A.T.
		#region Busca
		private void btBuscarArticulos(object sender, System.EventArgs e)
		{
			string stBusca = txtBusca.Text;
			string stExec = string.Format("Exec VentaArticuloLista '{0}'", stBusca);
			ugrBusca.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);
		}

		private void txtBusca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar != 13) return;
			string stBusca = txtBusca.Text.Trim();
			string stSelect = "Select Count(*) FROM " +
				"Articulo Where Codigo like '%" + stBusca + "%'" +
				" Or CodAlterno like '%" + stBusca + "%'" +
				" Or Articulo like '%" + stBusca + "%'";
			int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
			if (iCuenta == 1)
			{
				stSelect = "Select idArticulo FROM " +
					"Articulo Where Codigo like '%" + stBusca + "%'" +
					" Or CodAlterno like '%" + stBusca + "%'" +
					" Or Articulo like '%" + stBusca + "%'";
				int idArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);

				bool bExisteArt = false;
				// Recorre grilla para buscar si ya fue ingresado
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						if ((int) dr.Cells["idArticulo"].Value == idArt)
						{
							dr.Cells["Cantidad"].Value = (double) dr.Cells["Cantidad"].Value + 1;
							if (!chkCodBarras.Checked) dr.Activate();
							bExisteArt = true;
							break;
						}
					}
				}
				// Si no existe crea linea nueva en caso contrario solo aumenta la cantidad
				if (!bExisteArt)
				{
					ultraGrid1.Rows.Band.AddNew();
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
					ultraGrid1.ActiveCell.Value = idArt;
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
					ultraGrid1.ActiveCell.Value = idArt;
				}
				if (chkCodBarras.Checked)	txtBusca.SelectAll();
				else 
				{
					ultraGrid1.Select();
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
				}
			}
			else if (iCuenta > 1)
			{
				btBuscarArticulos(this, e);
			}
			else
				MessageBox.Show("No existen articulos que cumplan condicion", "Información");
		}

		private void btIngresar_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (ugrBusca.ActiveRow == null) return;
			int idArt = (int) ugrBusca.ActiveRow.Cells["idArticulo"].Value;

			ultraGrid1.Rows.Band.AddNew();
			ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
			ultraGrid1.ActiveCell.Value = idArt;
			if (drSeteoF["VerDesplegable"] != DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				string Codigo = ugrBusca.ActiveRow.Cells["Codigo"].Value.ToString();
				string Articulo = ugrBusca.ActiveRow.Cells["Nombre"].Value.ToString();
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Codigo"];
				ultraGrid1.ActiveCell.Value = Codigo;
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Articulo"];
				ultraGrid1.ActiveCell.Value = Articulo;
			}
			else
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
				ultraGrid1.ActiveCell.Value = idArt;
			}

			ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
		}

		#endregion Busca
		#region autoriz
		private void chkAutPrincipal_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (chkAutPrincipal.Checked) return;
			using (CompraAutoriza miAutoriza = new CompraAutoriza())
			{
				if (DialogResult.OK == miAutoriza.ShowDialog())
				{
					drCompra["AutPrincipal"] = true;
				}
				else
				{
					drCompra["AutPrincipal"] = false;
					chkAutPrincipal.Checked = false;
				}
			}
		}

		private void txtAutFactura_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			#region Verificaciones
			if (e.Button != MouseButtons.Right) return;
			if (IdTipoFactura != (int) Funcion.TipoFactura.Compra) return;
			if (drCompra["idCliente"] == null || (int) drCompra["idCliente"] == 0) 
			{
				MessageBox.Show("Selecione el Proveedor", "Información");
				return;
			}
			#endregion Verificaciones
			#region Liquidacion de Compras
			if ((int) drCompra["idComprobante"] == 3)
			{
				drCompra["AutFactura"] = drSeteo["AutLiquida"];
				drCompra["SerieFactura"] = drSeteo["SerieLiquida"];
				drCompra["FechaCaducidad"] = drSeteo["FechaCadLiq"];
				return;
			}
			#endregion
			#region Asignacion de datos
			drCompra["AutFactura"] = miProveedor.AutorizaFactura;
			drCompra["AutImprenta"] = miProveedor.AutorizaImprenta;
			drCompra["SerieFactura"] =miProveedor.SerieFactura;
			if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-10))
				drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
			#endregion Asignacion de datos
		}

		#endregion autoriza
		#region Otro
		private void cmbRecibeFabara_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbRecibeFabara.Value = 0;
			cmbRecibeFabara.Select();
			cmbRecibeFabara.Refresh();
		}

		private void btnComisionExt_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar comisiones mientras edita", "Información");
				return;
			}

			using (ConfitecaIngresaPC	miComisionC = new ConfitecaIngresaPC())
			{
				miComisionC.ShowDialog();
			}
			cdsConfiteca.Clear();
			cdsConfiteca.Fill();
		}
		#endregion Otro
		#endregion Pestanas

		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" 
				+ stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}

		private void controlCCP()
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRCentroC '" + txtE + "',1";
					iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from Proyecto where idProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa, true);							
						controlW = CentroC;							
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
						using (BuscaCentroCosto miBusca = new BuscaCentroCosto(txtE))
						{
							if (DialogResult.OK == miBusca.ShowDialog())
							{
								try
								{
									int id = miBusca.idBusca;
									string stEmpresa1 = "Select Nombre from Proyecto where idProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa1, true);									
									controlW = CentroC1;									
									ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}
		private void controlCPP(string cod1)
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRProyecto '" + txtE + "',1,"+cod1;
					iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from SubProyecto where idSubProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa, true);							
						controlW1 = CentroC;							
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
					
						using (BuscaCentroCosto miBusca1 = new BuscaCentroCosto(txtE,cod1))
						{
							if (DialogResult.OK == miBusca1.ShowDialog())
							{
								try
								{
									int id = miBusca1.idBusca;
									string stEmpresa1 = "Select Nombre from SubProyecto where idSubProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa1, true);									
									controlW1 = CentroC1;									
									ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}

		private string codCC(string cod)
		{
			string stCodCC = "Select CentroCosto from Proyecto where nombre like '%"+cod+"%'";
			string IdGlobal = Funcion.sEscalarSQL(cdsSeteoF, stCodCC, true);	
			return IdGlobal;
		}
		private void btnCopiarCotiz_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar comisiones mientras edita", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Copiar la Cotización?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			//			string stSelect = string.Format("Exec P_CopiarCotizacion {0},{1}",drCompra["idCompra"].ToString(),spnNumCopiarCotiz.Value);
			//			int control = Funcion.iEscalarSQL(cdsCompra,stSelect);
			int i=0;
			int idRetorna = 0;
			string stSelect="";
			while (int.Parse(spnNumCopiarCotiz.Value.ToString()) > i)
			{							
				idRetorna = CopiarFactura(3);				

				stSelect = string.Format("Exec P_CopiarCotizacion {0},{1},{2},{3}"
					,drCompra["idCompra"].ToString(),idRetorna,spnNumFrecCopiarCotiz.Value.ToString(),
					i);

				Funcion.EjecutaSQL(cdsCompra,stSelect);
				i++;
			}
			if (idRetorna > 1 )
				MessageBox.Show("Número de Cotización Copiadas... "+i );
			else
				MessageBox.Show("No se copio la Cotización...");
		}

		private void btDigitalizacion_Click(object sender, System.EventArgs e)
		{
			int	IdAsiento = (int) txtIdCompra.Value;
			int tipo=10;
			using (CIngresos miImagen = new CIngresos(IdAsiento,tipo))
			{
				miImagen.ShowDialog();
			}
		}

		private void btDigitalizacion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (ReporteDig miImagen = new ReporteDig())
			{
				miImagen.ShowDialog();
			}
		}

	}
}
