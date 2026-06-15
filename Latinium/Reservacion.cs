using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Reservacion.
	/// </summary>
	public class Reservacion : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdbComisionAgencias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOtroPorcentaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtServicios;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIntermediacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubcidio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBono;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private System.Windows.Forms.Panel panel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.Misc.UltraButton btGenerarFactura;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ckPorcentaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaIntermediacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtServiciosAdional;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalConIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalSinIva;
		private C1.Data.C1DataSet cdsReservacion;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPasajeros;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumPasajeros;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdReservacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReserva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotaImpuesto;
		private Infragistics.Win.Misc.UltraLabel ultraLabel18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraButton btReserva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor FormImpuesto;
		private Infragistics.Win.Misc.UltraButton btCalcImpuesto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor FormReserva;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor FormSubencion;
		private Infragistics.Win.Misc.UltraButton btSubencion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel19;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Label lblAnulado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroBono;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtImpuestos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtReservacion;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFee;
		private Infragistics.Win.Misc.UltraLabel ultraLabel20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		private System.Windows.Forms.Panel panel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredito2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContado2;
		private Infragistics.Win.Misc.UltraButton btCalculoFee;
		private Infragistics.Win.Misc.UltraButton btGeneraComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompraAux;
		private Infragistics.Win.Misc.UltraLabel ultraLabel21;
		private Infragistics.Win.Misc.UltraLabel ultraLabel22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContado1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredito1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button Copiar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Reservacion()
		{
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaSwift");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Beneficiario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoBarras");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NResolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetieneIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BloqueadoCrediPoint");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Courier");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProveedorIESS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PLNSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaBajoCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBancocliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("numCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigoBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaLineaCredito");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsReservacion = new C1.Data.C1DataSet();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.rdbComisionAgencias = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtOtroPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtServicios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIntermediacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaIntermediacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubcidio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBono = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.panel1 = new System.Windows.Forms.Panel();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel20 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva12 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtServiciosAdional = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalConIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalSinIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btGenerarFactura = new Infragistics.Win.Misc.UltraButton();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.ckPorcentaje = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNumPasajeros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtPasajeros = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdReservacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.txtReserva = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotaImpuesto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel18 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.FormReserva = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btReserva = new Infragistics.Win.Misc.UltraButton();
			this.FormImpuesto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCalcImpuesto = new Infragistics.Win.Misc.UltraButton();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.FormSubencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSubencion = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel19 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.txtNumeroBono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.txtImpuestos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtReservacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.txtFee = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCredito2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContado2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.panel2 = new System.Windows.Forms.Panel();
			this.ultraLabel22 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel21 = new Infragistics.Win.Misc.UltraLabel();
			this.btCalculoFee = new Infragistics.Win.Misc.UltraButton();
			this.btGeneraComprobante = new Infragistics.Win.Misc.UltraButton();
			this.txtIdCompraAux = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContado1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCredito1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.Copiar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsReservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdbComisionAgencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIntermediacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaIntermediacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubcidio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBono)).BeginInit();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServiciosAdional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalConIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalSinIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumPasajeros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajeros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdReservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaImpuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.FormReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.FormImpuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.FormSubencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroBono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImpuestos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFee)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado2)).BeginInit();
			this.panel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompraAux)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 19);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Cliente";
			// 
			// cdsReservacion
			// 
			this.cdsReservacion.BindingContextCtrl = this;
			this.cdsReservacion.DataLibrary = "LibFacturacion";
			this.cdsReservacion.DataLibraryUrl = "";
			this.cdsReservacion.DataSetDef = "dsReservacion";
			this.cdsReservacion.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsReservacion.Name = "cdsReservacion";
			this.cdsReservacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsReservacion.SchemaDef = null;
			this.cdsReservacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsReservacion_BeforeFill);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(8, 40);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(64, 14);
			this.ultraLabel2.TabIndex = 5;
			this.ultraLabel2.Text = "Reservacion";
			// 
			// rdbComisionAgencias
			// 
			this.rdbComisionAgencias.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance1.ForeColor = System.Drawing.Color.Black;
			this.rdbComisionAgencias.ItemAppearance = appearance1;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			appearance2.FontData.BoldAsString = "True";
			valueListItem1.Appearance = appearance2;
			valueListItem1.DataValue = 10;
			valueListItem1.DisplayText = "10 %";
			appearance3.FontData.BoldAsString = "True";
			valueListItem2.Appearance = appearance3;
			valueListItem2.DataValue = 12;
			valueListItem2.DisplayText = "12 %";
			appearance4.FontData.BoldAsString = "True";
			valueListItem3.Appearance = appearance4;
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "0 %";
			appearance5.FontData.BoldAsString = "True";
			valueListItem4.Appearance = appearance5;
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Otros %";
			this.rdbComisionAgencias.Items.Add(valueListItem1);
			this.rdbComisionAgencias.Items.Add(valueListItem2);
			this.rdbComisionAgencias.Items.Add(valueListItem3);
			this.rdbComisionAgencias.Items.Add(valueListItem4);
			this.rdbComisionAgencias.Location = new System.Drawing.Point(312, 176);
			this.rdbComisionAgencias.Name = "rdbComisionAgencias";
			this.rdbComisionAgencias.Size = new System.Drawing.Size(67, 62);
			this.rdbComisionAgencias.TabIndex = 12;
			this.rdbComisionAgencias.ValueChanged += new System.EventHandler(this.rdbComisionAgencias_ValueChanged);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(216, 176);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(90, 14);
			this.ultraLabel3.TabIndex = 7;
			this.ultraLabel3.Text = "Comision Agencia";
			// 
			// txtOtroPorcentaje
			// 
			this.txtOtroPorcentaje.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.PorComision"));
			this.txtOtroPorcentaje.Enabled = false;
			this.txtOtroPorcentaje.FormatString = "##0.00";
			this.txtOtroPorcentaje.Location = new System.Drawing.Point(440, 216);
			this.txtOtroPorcentaje.Name = "txtOtroPorcentaje";
			this.txtOtroPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtOtroPorcentaje.PromptChar = ' ';
			this.txtOtroPorcentaje.Size = new System.Drawing.Size(168, 21);
			this.txtOtroPorcentaje.TabIndex = 13;
			// 
			// txtServicios
			// 
			this.txtServicios.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Servicios"));
			this.txtServicios.FormatString = "#,##0.00";
			this.txtServicios.Location = new System.Drawing.Point(96, 240);
			this.txtServicios.Name = "txtServicios";
			this.txtServicios.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtServicios.PromptChar = ' ';
			this.txtServicios.Size = new System.Drawing.Size(96, 21);
			this.txtServicios.TabIndex = 8;
			// 
			// txtIntermediacion
			// 
			this.txtIntermediacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Intermediacion"));
			this.txtIntermediacion.Enabled = false;
			this.txtIntermediacion.FormatString = "#,##0.00";
			this.txtIntermediacion.Location = new System.Drawing.Point(512, 248);
			this.txtIntermediacion.Name = "txtIntermediacion";
			this.txtIntermediacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIntermediacion.PromptChar = ' ';
			this.txtIntermediacion.ReadOnly = true;
			this.txtIntermediacion.Size = new System.Drawing.Size(96, 21);
			this.txtIntermediacion.TabIndex = 12;
			// 
			// txtIvaIntermediacion
			// 
			this.txtIvaIntermediacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.IvaIntermediacion"));
			this.txtIvaIntermediacion.FormatString = "#,##0.00";
			this.txtIvaIntermediacion.Location = new System.Drawing.Point(328, 248);
			this.txtIvaIntermediacion.Name = "txtIvaIntermediacion";
			this.txtIvaIntermediacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaIntermediacion.PromptChar = ' ';
			this.txtIvaIntermediacion.Size = new System.Drawing.Size(94, 21);
			this.txtIvaIntermediacion.TabIndex = 14;
			// 
			// txtSubcidio
			// 
			this.txtSubcidio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Subsidio"));
			this.txtSubcidio.FormatString = "#,##0.00";
			this.txtSubcidio.Location = new System.Drawing.Point(472, 136);
			this.txtSubcidio.Name = "txtSubcidio";
			this.txtSubcidio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubcidio.PromptChar = ' ';
			this.txtSubcidio.ReadOnly = true;
			this.txtSubcidio.Size = new System.Drawing.Size(136, 21);
			this.txtSubcidio.TabIndex = 13;
			// 
			// txtBono
			// 
			this.txtBono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Bono"));
			this.txtBono.FormatString = "#,##0.00";
			this.txtBono.Location = new System.Drawing.Point(96, 168);
			this.txtBono.Name = "txtBono";
			this.txtBono.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBono.PromptChar = ' ';
			this.txtBono.Size = new System.Drawing.Size(96, 21);
			this.txtBono.TabIndex = 6;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(8, 168);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(29, 14);
			this.ultraLabel4.TabIndex = 15;
			this.ultraLabel4.Text = "Bono";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(8, 136);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(60, 14);
			this.ultraLabel5.TabIndex = 16;
			this.ultraLabel5.Text = "Subvención";
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(224, 248);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(91, 14);
			this.ultraLabel7.TabIndex = 18;
			this.ultraLabel7.Text = "Intermediacion Iva";
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Location = new System.Drawing.Point(8, 240);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(48, 14);
			this.ultraLabel8.TabIndex = 19;
			this.ultraLabel8.Text = "Servicios";
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.AutoSize = true;
			this.ultraLabel9.Location = new System.Drawing.Point(8, 88);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(53, 14);
			this.ultraLabel9.TabIndex = 20;
			this.ultraLabel9.Text = "Impuestos";
			// 
			// panel1
			// 
			this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel1.Controls.Add(this.txtDescuento);
			this.panel1.Controls.Add(this.ultraLabel20);
			this.panel1.Controls.Add(this.ultraLabel15);
			this.panel1.Controls.Add(this.ultraLabel13);
			this.panel1.Controls.Add(this.ultraLabel14);
			this.panel1.Controls.Add(this.ultraLabel12);
			this.panel1.Controls.Add(this.ultraLabel11);
			this.panel1.Controls.Add(this.txtFactura);
			this.panel1.Controls.Add(this.txtIva12);
			this.panel1.Controls.Add(this.txtServiciosAdional);
			this.panel1.Controls.Add(this.txtTotalConIva);
			this.panel1.Controls.Add(this.txtTotalSinIva);
			this.panel1.Location = new System.Drawing.Point(376, 280);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(232, 144);
			this.panel1.TabIndex = 27;
			// 
			// txtDescuento
			// 
			this.txtDescuento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Descuento"));
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(120, 72);
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(93, 21);
			this.txtDescuento.TabIndex = 39;
			// 
			// ultraLabel20
			// 
			this.ultraLabel20.AutoSize = true;
			this.ultraLabel20.Location = new System.Drawing.Point(24, 72);
			this.ultraLabel20.Name = "ultraLabel20";
			this.ultraLabel20.Size = new System.Drawing.Size(55, 14);
			this.ultraLabel20.TabIndex = 38;
			this.ultraLabel20.Text = "Descuento";
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.AutoSize = true;
			this.ultraLabel15.Font = new System.Drawing.Font("Arial", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel15.Location = new System.Drawing.Point(24, 120);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(71, 13);
			this.ultraLabel15.TabIndex = 37;
			this.ultraLabel15.Text = "Total Factura";
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.Location = new System.Drawing.Point(24, 96);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(45, 14);
			this.ultraLabel13.TabIndex = 36;
			this.ultraLabel13.Text = "Iva 12 %";
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.AutoSize = true;
			this.ultraLabel14.Location = new System.Drawing.Point(20, 51);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(94, 14);
			this.ultraLabel14.TabIndex = 35;
			this.ultraLabel14.Text = "Servicios Adicional";
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.AutoSize = true;
			this.ultraLabel12.Location = new System.Drawing.Point(20, 32);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(67, 14);
			this.ultraLabel12.TabIndex = 33;
			this.ultraLabel12.Text = "Total Con Iva";
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.AutoSize = true;
			this.ultraLabel11.Location = new System.Drawing.Point(20, 13);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(63, 14);
			this.ultraLabel11.TabIndex = 32;
			this.ultraLabel11.Text = "Total Sin Iva";
			// 
			// txtFactura
			// 
			this.txtFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.TotalFactura"));
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtFactura.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFactura.FormatString = "#,##0.00";
			this.txtFactura.Location = new System.Drawing.Point(120, 112);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFactura.PromptChar = ' ';
			this.txtFactura.ReadOnly = true;
			this.txtFactura.Size = new System.Drawing.Size(93, 22);
			this.txtFactura.TabIndex = 4;
			// 
			// txtIva12
			// 
			this.txtIva12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.TotalIva"));
			this.txtIva12.FormatString = "#,##0.00";
			this.txtIva12.Location = new System.Drawing.Point(120, 96);
			this.txtIva12.Name = "txtIva12";
			this.txtIva12.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva12.PromptChar = ' ';
			this.txtIva12.ReadOnly = true;
			this.txtIva12.Size = new System.Drawing.Size(93, 21);
			this.txtIva12.TabIndex = 3;
			// 
			// txtServiciosAdional
			// 
			this.txtServiciosAdional.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.TotalServicios"));
			this.txtServiciosAdional.FormatString = "#,##0.00";
			this.txtServiciosAdional.Location = new System.Drawing.Point(120, 48);
			this.txtServiciosAdional.Name = "txtServiciosAdional";
			this.txtServiciosAdional.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtServiciosAdional.PromptChar = ' ';
			this.txtServiciosAdional.ReadOnly = true;
			this.txtServiciosAdional.Size = new System.Drawing.Size(93, 21);
			this.txtServiciosAdional.TabIndex = 2;
			// 
			// txtTotalConIva
			// 
			this.txtTotalConIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.TotalConIva"));
			this.txtTotalConIva.FormatString = "#,##0.00";
			this.txtTotalConIva.Location = new System.Drawing.Point(120, 30);
			this.txtTotalConIva.Name = "txtTotalConIva";
			this.txtTotalConIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalConIva.PromptChar = ' ';
			this.txtTotalConIva.ReadOnly = true;
			this.txtTotalConIva.Size = new System.Drawing.Size(93, 21);
			this.txtTotalConIva.TabIndex = 1;
			// 
			// txtTotalSinIva
			// 
			this.txtTotalSinIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.TotalSinIva"));
			this.txtTotalSinIva.FormatString = "#,##0.00";
			this.txtTotalSinIva.Location = new System.Drawing.Point(120, 11);
			this.txtTotalSinIva.Name = "txtTotalSinIva";
			this.txtTotalSinIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalSinIva.PromptChar = ' ';
			this.txtTotalSinIva.ReadOnly = true;
			this.txtTotalSinIva.Size = new System.Drawing.Size(93, 21);
			this.txtTotalSinIva.TabIndex = 0;
			// 
			// btGenerarFactura
			// 
			this.btGenerarFactura.Location = new System.Drawing.Point(360, 432);
			this.btGenerarFactura.Name = "btGenerarFactura";
			this.btGenerarFactura.Size = new System.Drawing.Size(60, 24);
			this.btGenerarFactura.TabIndex = 22;
			this.btGenerarFactura.Text = "Generar";
			this.btGenerarFactura.Click += new System.EventHandler(this.btGenerarFactura_Click);
			// 
			// txtComision
			// 
			this.txtComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Comision"));
			this.txtComision.Enabled = false;
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(472, 168);
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(136, 21);
			this.txtComision.TabIndex = 18;
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.AutoSize = true;
			this.ultraLabel16.Location = new System.Drawing.Point(384, 168);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(29, 14);
			this.ultraLabel16.TabIndex = 31;
			this.ultraLabel16.Text = "Valor";
			// 
			// ckPorcentaje
			// 
			this.ckPorcentaje.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsReservacion, "Reservacion.Porcentaje"));
			this.ckPorcentaje.Location = new System.Drawing.Point(448, 248);
			this.ckPorcentaje.Name = "ckPorcentaje";
			this.ckPorcentaje.Size = new System.Drawing.Size(34, 16);
			this.ckPorcentaje.TabIndex = 32;
			this.ckPorcentaje.Text = "%";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			this.cmbCliente.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Articulo";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 8;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 43;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 50;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 17;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 17;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 17;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 8;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 8;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 17;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 8;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 8;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 8;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 8;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 17;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 17;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 8;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 8;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 8;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 8;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Width = 8;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 8;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 17;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 17;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Width = 8;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 8;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 17;
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 8;
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Width = 17;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Width = 17;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Width = 17;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Width = 8;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Width = 17;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Width = 17;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Width = 17;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Width = 8;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.Width = 8;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Width = 17;
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
																										 ultraGridColumn55});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(96, 16);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(240, 21);
			this.cmbCliente.TabIndex = 0;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "Reservacion";
			this.barraDatoSQL1.DataNombreId = "idReservacion";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsReservacion;
			this.barraDatoSQL1.DataTabla = "Reservacion";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(24, 432);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 15;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtNumPasajeros
			// 
			this.txtNumPasajeros.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.NumBoletos"));
			this.txtNumPasajeros.Location = new System.Drawing.Point(96, 264);
			this.txtNumPasajeros.Name = "txtNumPasajeros";
			this.txtNumPasajeros.PromptChar = ' ';
			this.txtNumPasajeros.Size = new System.Drawing.Size(96, 21);
			this.txtNumPasajeros.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumPasajeros.TabIndex = 9;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 264);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 38;
			this.label1.Text = "No. Pasajeros";
			// 
			// txtPasajeros
			// 
			this.txtPasajeros.AcceptsReturn = true;
			this.txtPasajeros.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Pasajeros"));
			this.txtPasajeros.Location = new System.Drawing.Point(200, 344);
			this.txtPasajeros.Multiline = true;
			this.txtPasajeros.Name = "txtPasajeros";
			this.txtPasajeros.Size = new System.Drawing.Size(168, 80);
			this.txtPasajeros.TabIndex = 16;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.idFactura"));
			this.txtIdCompra.FormatString = "#,##0.00";
			this.txtIdCompra.Location = new System.Drawing.Point(400, 8);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.Nullable = true;
			this.txtIdCompra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 40;
			// 
			// txtIdReservacion
			// 
			this.txtIdReservacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.idReservacion"));
			this.txtIdReservacion.FormatString = "#,##0.00";
			this.txtIdReservacion.Location = new System.Drawing.Point(344, 8);
			this.txtIdReservacion.Name = "txtIdReservacion";
			this.txtIdReservacion.Nullable = true;
			this.txtIdReservacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdReservacion.PromptChar = ' ';
			this.txtIdReservacion.Size = new System.Drawing.Size(48, 21);
			this.txtIdReservacion.TabIndex = 41;
			// 
			// txtNumero
			// 
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Numero"));
			this.txtNumero.FormatString = "";
			this.txtNumero.Location = new System.Drawing.Point(512, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Nullable = true;
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.ReadOnly = true;
			this.txtNumero.Size = new System.Drawing.Size(96, 21);
			this.txtNumero.TabIndex = 42;
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(328, 432);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(24, 23);
			this.btFactura.TabIndex = 23;
			this.btFactura.Text = "F";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// txtReserva
			// 
			this.txtReserva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Reserva"));
			this.txtReserva.Location = new System.Drawing.Point(96, 64);
			this.txtReserva.Name = "txtReserva";
			this.txtReserva.Size = new System.Drawing.Size(512, 21);
			this.txtReserva.TabIndex = 4;
			// 
			// txtNotaImpuesto
			// 
			this.txtNotaImpuesto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Impuest"));
			this.txtNotaImpuesto.Location = new System.Drawing.Point(96, 112);
			this.txtNotaImpuesto.Name = "txtNotaImpuesto";
			this.txtNotaImpuesto.Size = new System.Drawing.Size(512, 21);
			this.txtNotaImpuesto.TabIndex = 4;
			// 
			// ultraLabel18
			// 
			this.ultraLabel18.AutoSize = true;
			this.ultraLabel18.Location = new System.Drawing.Point(8, 192);
			this.ultraLabel18.Name = "ultraLabel18";
			this.ultraLabel18.Size = new System.Drawing.Size(70, 14);
			this.ultraLabel18.TabIndex = 51;
			this.ultraLabel18.Text = "Nombre Bono";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.NombreBono"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(96, 192);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor1.TabIndex = 7;
			// 
			// FormReserva
			// 
			this.FormReserva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.FormReserva"));
			this.FormReserva.Location = new System.Drawing.Point(96, 40);
			this.FormReserva.Name = "FormReserva";
			this.FormReserva.Size = new System.Drawing.Size(240, 21);
			this.FormReserva.TabIndex = 1;
			this.FormReserva.Validated += new System.EventHandler(this.FormReserva_Validated);
			// 
			// btReserva
			// 
			this.btReserva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btReserva.Location = new System.Drawing.Point(392, 40);
			this.btReserva.Name = "btReserva";
			this.btReserva.Size = new System.Drawing.Size(40, 23);
			this.btReserva.TabIndex = 2;
			this.btReserva.Text = "=>";
			this.btReserva.Click += new System.EventHandler(this.btReserva_Click);
			// 
			// FormImpuesto
			// 
			this.FormImpuesto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.FormImpuesto"));
			this.FormImpuesto.Location = new System.Drawing.Point(96, 88);
			this.FormImpuesto.Name = "FormImpuesto";
			this.FormImpuesto.Size = new System.Drawing.Size(240, 21);
			this.FormImpuesto.TabIndex = 5;
			this.FormImpuesto.Validated += new System.EventHandler(this.FormImpuesto_Validated);
			// 
			// btCalcImpuesto
			// 
			this.btCalcImpuesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btCalcImpuesto.Location = new System.Drawing.Point(392, 88);
			this.btCalcImpuesto.Name = "btCalcImpuesto";
			this.btCalcImpuesto.Size = new System.Drawing.Size(40, 23);
			this.btCalcImpuesto.TabIndex = 6;
			this.btCalcImpuesto.Text = "=>";
			this.btCalcImpuesto.Click += new System.EventHandler(this.btCalcImpuesto_Click);
			// 
			// ultraCalendarCombo1
			// 
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.FechaSalida"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(256, 320);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 21;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(112, 21);
			this.ultraCalendarCombo1.SpinButtonsVisible = true;
			this.ultraCalendarCombo1.TabIndex = 11;
			this.ultraCalendarCombo1.Value = new System.DateTime(2005, 12, 19, 0, 0, 0, 0);
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.Location = new System.Drawing.Point(200, 320);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(46, 14);
			this.ultraLabel10.TabIndex = 53;
			this.ultraLabel10.Text = "F. Salida";
			// 
			// FormSubencion
			// 
			this.FormSubencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.FormSubsidio"));
			this.FormSubencion.Location = new System.Drawing.Point(96, 136);
			this.FormSubencion.Name = "FormSubencion";
			this.FormSubencion.Size = new System.Drawing.Size(240, 21);
			this.FormSubencion.TabIndex = 5;
			this.FormSubencion.Validated += new System.EventHandler(this.FormSubencion_Validated);
			// 
			// btSubencion
			// 
			this.btSubencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btSubencion.Location = new System.Drawing.Point(392, 136);
			this.btSubencion.Name = "btSubencion";
			this.btSubencion.Size = new System.Drawing.Size(40, 23);
			this.btSubencion.TabIndex = 10;
			this.btSubencion.Text = "=>";
			this.btSubencion.Click += new System.EventHandler(this.btSubencion_Click);
			// 
			// cmbDesde
			// 
			this.cmbDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Fecha"));
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(256, 296);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 21;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 54;
			this.cmbDesde.Value = new System.DateTime(2005, 12, 19, 0, 0, 0, 0);
			// 
			// ultraLabel19
			// 
			this.ultraLabel19.AutoSize = true;
			this.ultraLabel19.Location = new System.Drawing.Point(200, 296);
			this.ultraLabel19.Name = "ultraLabel19";
			this.ultraLabel19.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel19.TabIndex = 55;
			this.ultraLabel19.Text = "Fecha";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// lblAnulado
			// 
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(432, 432);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(112, 23);
			this.lblAnulado.TabIndex = 56;
			this.lblAnulado.Text = "ANULADO";
			this.lblAnulado.Visible = false;
			// 
			// txtNumeroBono
			// 
			this.txtNumeroBono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.NumeroBono"));
			this.txtNumeroBono.Location = new System.Drawing.Point(96, 216);
			this.txtNumeroBono.Name = "txtNumeroBono";
			this.txtNumeroBono.Size = new System.Drawing.Size(96, 21);
			this.txtNumeroBono.TabIndex = 57;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(8, 216);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(70, 14);
			this.ultraLabel6.TabIndex = 58;
			this.ultraLabel6.Text = "Numero Bono";
			// 
			// txtImpuestos
			// 
			this.txtImpuestos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Impuestos"));
			this.txtImpuestos.FormatString = "#,##0.00";
			this.txtImpuestos.Location = new System.Drawing.Point(464, 88);
			this.txtImpuestos.Name = "txtImpuestos";
			this.txtImpuestos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtImpuestos.PromptChar = ' ';
			this.txtImpuestos.Size = new System.Drawing.Size(144, 21);
			this.txtImpuestos.TabIndex = 7;
			// 
			// txtReservacion
			// 
			this.txtReservacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Reservacion"));
			this.txtReservacion.FormatString = "#,##0.00";
			this.txtReservacion.Location = new System.Drawing.Point(464, 40);
			this.txtReservacion.Name = "txtReservacion";
			this.txtReservacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtReservacion.PromptChar = ' ';
			this.txtReservacion.Size = new System.Drawing.Size(144, 21);
			this.txtReservacion.TabIndex = 3;
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.AutoSize = true;
			this.ultraLabel17.Location = new System.Drawing.Point(384, 192);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(22, 14);
			this.ultraLabel17.TabIndex = 60;
			this.ultraLabel17.Text = "Fee";
			// 
			// txtFee
			// 
			this.txtFee.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.Fee"));
			this.txtFee.Enabled = false;
			this.txtFee.FormatString = "#,##0.00";
			this.txtFee.Location = new System.Drawing.Point(472, 192);
			this.txtFee.Name = "txtFee";
			this.txtFee.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFee.PromptChar = ' ';
			this.txtFee.Size = new System.Drawing.Size(136, 21);
			this.txtFee.TabIndex = 59;
			// 
			// txtCredito2
			// 
			this.txtCredito2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.VCreditoContab"));
			this.txtCredito2.FormatString = "#,##0.00";
			this.txtCredito2.Location = new System.Drawing.Point(88, 32);
			this.txtCredito2.Name = "txtCredito2";
			this.txtCredito2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredito2.PromptChar = ' ';
			this.txtCredito2.Size = new System.Drawing.Size(88, 21);
			this.txtCredito2.TabIndex = 64;
			// 
			// txtContado2
			// 
			this.txtContado2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.VContadoContab"));
			this.txtContado2.FormatString = "#,##0.00";
			this.txtContado2.Location = new System.Drawing.Point(88, 8);
			this.txtContado2.Name = "txtContado2";
			this.txtContado2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtContado2.PromptChar = ' ';
			this.txtContado2.Size = new System.Drawing.Size(88, 21);
			this.txtContado2.TabIndex = 63;
			// 
			// panel2
			// 
			this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel2.Controls.Add(this.ultraLabel22);
			this.panel2.Controls.Add(this.ultraLabel21);
			this.panel2.Controls.Add(this.txtCredito2);
			this.panel2.Controls.Add(this.txtContado2);
			this.panel2.Enabled = false;
			this.panel2.Location = new System.Drawing.Point(8, 360);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(184, 64);
			this.panel2.TabIndex = 65;
			this.panel2.Visible = false;
			// 
			// ultraLabel22
			// 
			this.ultraLabel22.AutoSize = true;
			this.ultraLabel22.Location = new System.Drawing.Point(8, 40);
			this.ultraLabel22.Name = "ultraLabel22";
			this.ultraLabel22.Size = new System.Drawing.Size(39, 14);
			this.ultraLabel22.TabIndex = 66;
			this.ultraLabel22.Text = "Credito";
			// 
			// ultraLabel21
			// 
			this.ultraLabel21.AutoSize = true;
			this.ultraLabel21.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel21.Name = "ultraLabel21";
			this.ultraLabel21.Size = new System.Drawing.Size(44, 14);
			this.ultraLabel21.TabIndex = 65;
			this.ultraLabel21.Text = "Contado";
			// 
			// btCalculoFee
			// 
			this.btCalculoFee.Location = new System.Drawing.Point(544, 432);
			this.btCalculoFee.Name = "btCalculoFee";
			this.btCalculoFee.Size = new System.Drawing.Size(48, 32);
			this.btCalculoFee.TabIndex = 67;
			this.btCalculoFee.Text = "Calc Fee";
			this.btCalculoFee.Click += new System.EventHandler(this.btCalculoFee_Click);
			// 
			// btGeneraComprobante
			// 
			this.btGeneraComprobante.Location = new System.Drawing.Point(600, 432);
			this.btGeneraComprobante.Name = "btGeneraComprobante";
			this.btGeneraComprobante.Size = new System.Drawing.Size(64, 32);
			this.btGeneraComprobante.TabIndex = 68;
			this.btGeneraComprobante.Text = "Gen Add";
			this.btGeneraComprobante.Visible = false;
			this.btGeneraComprobante.WrapText = false;
			this.btGeneraComprobante.Click += new System.EventHandler(this.btGeneraComprobante_Click);
			this.btGeneraComprobante.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btGeneraComprobante_MouseDown);
			// 
			// txtIdCompraAux
			// 
			this.txtIdCompraAux.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.idFacturaAux"));
			this.txtIdCompraAux.FormatString = "#,##0.00";
			this.txtIdCompraAux.Location = new System.Drawing.Point(520, 8);
			this.txtIdCompraAux.Name = "txtIdCompraAux";
			this.txtIdCompraAux.Nullable = true;
			this.txtIdCompraAux.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIdCompraAux.PromptChar = ' ';
			this.txtIdCompraAux.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompraAux.TabIndex = 69;
			// 
			// txtContado1
			// 
			this.txtContado1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.VContadoReserva"));
			this.txtContado1.FormatString = "#,##0.00";
			this.txtContado1.Location = new System.Drawing.Point(8, 320);
			this.txtContado1.Name = "txtContado1";
			this.txtContado1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtContado1.PromptChar = ' ';
			this.txtContado1.Size = new System.Drawing.Size(88, 21);
			this.txtContado1.TabIndex = 71;
			// 
			// txtCredito1
			// 
			this.txtCredito1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReservacion, "Reservacion.VCreditoReserva"));
			this.txtCredito1.FormatString = "#,##0.00";
			this.txtCredito1.Location = new System.Drawing.Point(104, 320);
			this.txtCredito1.Name = "txtCredito1";
			this.txtCredito1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredito1.PromptChar = ' ';
			this.txtCredito1.Size = new System.Drawing.Size(88, 21);
			this.txtCredito1.TabIndex = 72;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(24, 304);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 73;
			this.label2.Text = "Contado";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(112, 304);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 74;
			this.label3.Text = "Credito";
			// 
			// Copiar
			// 
			this.Copiar.Location = new System.Drawing.Point(616, 40);
			this.Copiar.Name = "Copiar";
			this.Copiar.Size = new System.Drawing.Size(24, 24);
			this.Copiar.TabIndex = 75;
			this.Copiar.Text = "...";
			this.Copiar.Click += new System.EventHandler(this.Copiar_Click);
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
			// 
			// Reservacion
			// 
			this.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Appearance.Options.UseFont = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
			this.ClientSize = new System.Drawing.Size(664, 469);
			this.Controls.Add(this.Copiar);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtCredito1);
			this.Controls.Add(this.txtContado1);
			this.Controls.Add(this.btGeneraComprobante);
			this.Controls.Add(this.btCalculoFee);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.ultraLabel17);
			this.Controls.Add(this.txtFee);
			this.Controls.Add(this.txtReservacion);
			this.Controls.Add(this.txtImpuestos);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.txtNumeroBono);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel19);
			this.Controls.Add(this.btSubencion);
			this.Controls.Add(this.FormSubencion);
			this.Controls.Add(this.ultraCalendarCombo1);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.btCalcImpuesto);
			this.Controls.Add(this.FormImpuesto);
			this.Controls.Add(this.btReserva);
			this.Controls.Add(this.FormReserva);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.ultraLabel18);
			this.Controls.Add(this.txtNotaImpuesto);
			this.Controls.Add(this.txtReserva);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtIdReservacion);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtPasajeros);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumPasajeros);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.ckPorcentaje);
			this.Controls.Add(this.ultraLabel16);
			this.Controls.Add(this.txtComision);
			this.Controls.Add(this.btGenerarFactura);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBono);
			this.Controls.Add(this.txtSubcidio);
			this.Controls.Add(this.txtIvaIntermediacion);
			this.Controls.Add(this.txtIntermediacion);
			this.Controls.Add(this.txtServicios);
			this.Controls.Add(this.txtOtroPorcentaje);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.rdbComisionAgencias);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtIdCompraAux);
			this.MaximizeBox = false;
			this.Name = "Reservacion";
			this.Text = "Agencia de Viajes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Reservacion_Closing);
			this.Load += new System.EventHandler(this.Reservacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsReservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdbComisionAgencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIntermediacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaIntermediacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubcidio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBono)).EndInit();
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServiciosAdional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalConIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalSinIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumPasajeros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajeros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdReservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotaImpuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.FormReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.FormImpuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.FormSubencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroBono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImpuestos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFee)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado2)).EndInit();
			this.panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompraAux)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular Reservaciones",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea Eliminar la Reservacion?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) ) 
				return;

			if (lblAnulado.Visible)
			{
				barraDatoSQL1.BorraRegistro();
				MessageBox.Show("Registro Eliminado Permanentemente", "Informacion");
			}
			else
			{
				drReserva["Borrar"] = true;
				cdsReservacion.Update();
				lblAnulado.Visible = true;
			}
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaReserva miBusqueda = new BuscaReserva())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					int id = 0;
					try
					{
						id = int.Parse(miBusqueda.txtIdBusca.Value.ToString());
					}
					catch {};
					if (id > 0)
					{
						barraDatoSQL1.IdRegistro = id;
						barraDatoSQL1.ProximoId(5);
					}
				}
			}
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
			ReservaC miReserva = new ReservaC();
			miReserva.MdiParent = this.MdiParent;
			miReserva.Show();
		}

		private void ValoresIniciales()
		{
			drReserva["Fecha"] = DateTime.Today;	
			drReserva["NumBoletos"] = 0;
			cmbCliente.Select();
			drReserva["VContadoContab"] =0;
			drReserva["VContadoReserva"] =0;
			drReserva["VCreditoReserva"] =0;
			drReserva["VCreditoContab"] =0;

		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a crear Reservaciones",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDatoSQL1.CrearRegistro();
			
			ValoresIniciales();
			rdbComisionAgencias_ValueChanged(this, e);
			panel2.Enabled = true;


		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
			panel2.Enabled = false;

		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Reservaciones",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDatoSQL1.EditarRegistro();
			rdbComisionAgencias_ValueChanged(this, e);
			panel2.Enabled = true;


		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			ReservaReporte miRep = new ReservaReporte();
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			bool bNuevo = barraDatoSQL1.bNuevo;
			
			if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FEEAUT'")) 				
			{
			if (CaluloFee()>0) return;
			}
			

			this.CalcularFactura();
			barraDatoSQL1.GrabaRegistro();
			if (bNuevo)
			{
				string stNumero1 = "";
				stNumero1 = Numera();
				txtNumero.Value = stNumero1;
				string stSelect = "Update Reservacion SET Numero = " + stNumero1
					+ " Where idReservacion = " + txtIdReservacion.Value.ToString();
				Funcion.EjecutaSQL(cdsReservacion, stSelect);
			}
			barraDatoSQL1.ProximoId(5);
//			btGenerarFactura_Click(this, e);
			panel2.Enabled = false;
	

		}

		private C1.Data.C1DataRow drReserva;
		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				if (txtOtroPorcentaje.Value == System.DBNull.Value) return;
				drReserva = cdsReservacion.TableViews["Reservacion"].Rows[0];
				if (int.Parse(txtOtroPorcentaje.Value.ToString()) == 10
				 || int.Parse(txtOtroPorcentaje.Value.ToString()) == 12 
				 || int.Parse(txtOtroPorcentaje.Value.ToString()) == 0 )
					rdbComisionAgencias.Value = int.Parse(txtOtroPorcentaje.Value.ToString());
				else
					rdbComisionAgencias.Value = 1;
				if (drReserva["Borrar"] != System.DBNull.Value)
					lblAnulado.Visible = (bool) drReserva["Borrar"];
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Refresca Barra Dato");
			} 
			
		}

		private void cdsReservacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsReservacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private string Numera()
		{
			string stNumero = "Select Numero FROM ReservaC";
			int iNumDespues = Funcion.iEscalarSQL(cdsReservacion, stNumero);
			stNumero = iNumDespues.ToString();
			iNumDespues ++;
			string stSelect = "Update ReservaC SET Numero = " + iNumDespues.ToString();
			Funcion.EjecutaSQL(cdsReservacion, stSelect);
			return stNumero;
		}

		C1.Data.C1DataRow drConstante;
		private Acceso miAcceso;
		private void Reservacion_Load(object sender, System.EventArgs e)
		{
		//seguridad
			//revisar si esta activa en el menu el visto en acceso

			ToolTip toolTip1 = new ToolTip();

			// Set up the delays for the ToolTip.
			toolTip1.AutoPopDelay = 5000;
			toolTip1.InitialDelay = 1000;
			toolTip1.ReshowDelay = 500;
			// Force the ToolTip text to be displayed whether or not the form is active.
			toolTip1.ShowAlways = true;
     
			// Set up the ToolTip text for the Button and Checkbox.
			toolTip1.SetToolTip(this.Copiar, "Copiar Suma de Reservacion e Impuesto a Casilla de Credito..");		

			bool acceso = Funcion.Permiso("725",cdsCliente);
			if (acceso == true)
			{
				panel2.Visible = true;
				btGeneraComprobante.Visible = true;

			}

			string stAudita = "Exec AuditaCrear 53, 6, 'Reservacion'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			miAcceso =new Acceso("0403", this.cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			try
			{
				barraDatoSQL1.PosUltima();
			}
			catch{};
			barraDatoSQL1.HabilitaControles(false);
			txtIdCompra.Width = 0;
			txtIdReservacion.Width = 0;
//			btGenerarFactura.Width = 0;
			// Si no tiene acceso a editar facturas de venta no puede cambiar porcentaje
			if (!Funcion.Permiso("20", cdsCliente))
				txtOtroPorcentaje.ReadOnly = true;
			drConstante = cdsReservacion.TableViews["ReservaC"].Rows[0];
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtReservacion, "");
			errorProvider.SetError(rdbComisionAgencias, "");
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(txtNumPasajeros, "");
			errorProvider.SetError(ultraCalendarCombo1, "");
				errorProvider.SetError(txtContado1, "");
				errorProvider.SetError(txtCredito1, "");
			
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (double.Parse(txtReservacion.Value.ToString()) <= 0.0)
			{
				errorProvider.SetError(txtReservacion, "Ingrese Reservacion");
				bOk = false;
			}
			if (rdbComisionAgencias.Value == null
				|| (int) rdbComisionAgencias.Value < 0)
			{
				errorProvider.SetError(rdbComisionAgencias, "Ingrese % de Comision");
				bOk = false;
			}
			if (cmbCliente.Value == System.DBNull.Value)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				bOk = false;
			}
			if (txtNumPasajeros.Value == System.DBNull.Value || (int) txtNumPasajeros.Value <= 0)
			{
				errorProvider.SetError(txtNumPasajeros, "Ingrese No. de Pasajeros");
				bOk = false;
			}
			if (ultraCalendarCombo1.Value == System.DBNull.Value || (DateTime) ultraCalendarCombo1.Value < (DateTime)cmbDesde.Value)
			{
				errorProvider.SetError(ultraCalendarCombo1, "Ingrese Fecha de Salida");
				bOk = false;
			}
			if((double)txtCredito1.Value == 0.0 && (double)txtContado1.Value == 0.0)
			{
				errorProvider.SetError(txtCredito1, "Ingrese Valor Credito");
				errorProvider.SetError(txtContado1, "Ingrese Valor Contado");
				bOk = false;

			}
					 return bOk;
		}

		private void CalcularFactura()
		{
			if (!Verifica()) return;
			float lnReservacion,lnPorComision;
			float lnComision,lnBono,lnIntenmediacion,lnIva12;
			float lnIvaIntemediacion,lnImpuestos,lnServicios,lnSubsidio;
			float lnTotalSinIva,lnServiciosAdi, lnFee, lnDescuento;

			lnReservacion=float.Parse(txtReservacion.Value.ToString());
			lnBono=float.Parse(txtBono.Value.ToString());
			lnSubsidio=float.Parse(txtSubcidio.Value.ToString());
			lnServicios=float.Parse(txtServicios.Value.ToString());
			lnImpuestos=float.Parse(txtImpuestos.Value.ToString());
			lnFee=float.Parse(txtFee.Value.ToString());
			lnDescuento=float.Parse(txtDescuento.Value.ToString());

       if ((int)rdbComisionAgencias.Value==1)
				lnPorComision =  float.Parse (txtOtroPorcentaje.Value.ToString()) ;
			else
				lnPorComision =  float.Parse (rdbComisionAgencias.Value.ToString()) ;

			lnComision=(float) lnReservacion*(lnPorComision/100);
			txtComision.Value=lnComision;
			
			if (ckPorcentaje.Checked==true)
			{
				lnIvaIntemediacion=float.Parse(txtIvaIntermediacion.Value.ToString())/0.12f;
				lnIvaIntemediacion=float.Parse(txtReservacion.Value.ToString()) *(lnIvaIntemediacion/100);
			}
			else
				lnIvaIntemediacion=float.Parse(txtIvaIntermediacion.Value.ToString())/0.12f;
			
			lnIva12=(lnIvaIntemediacion + lnFee - lnDescuento) *0.12f;
			txtIntermediacion.Value=lnIvaIntemediacion;
			lnIntenmediacion=lnIvaIntemediacion + lnFee - lnDescuento;
			//Calculo el Valor Total Sin Iva
			lnTotalSinIva = lnReservacion -lnComision -lnIntenmediacion
				+lnFee-lnDescuento;
//				- lnImpuestos - lnServicios;
			txtTotalSinIva.Value=lnTotalSinIva;
			// Calculo el Valor Total con Iva
			txtTotalConIva.Value=lnIntenmediacion;
			// Calculo o sumo todos los Adicionales
			lnServiciosAdi=lnImpuestos+lnServicios;
			txtServiciosAdional.Value=lnServiciosAdi;

			txtIva12.Value=lnIva12;
			txtFactura.Value=lnTotalSinIva+lnIntenmediacion+lnServiciosAdi+lnIva12;
		}

		

		private void rdbComisionAgencias_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)rdbComisionAgencias.Value != 1) 
			{
				txtOtroPorcentaje.Value=float.Parse (rdbComisionAgencias.Value.ToString());
				txtOtroPorcentaje.Select();
				txtOtroPorcentaje.Refresh();
			}

			if ((int)rdbComisionAgencias.Value == 1)
				if (barraDatoSQL1.bEditar) txtOtroPorcentaje.Enabled=true;
				else txtOtroPorcentaje.Enabled=false;
			else
				txtOtroPorcentaje.Enabled=false;
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btGenerarFactura_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Generar Factura?",
				"Confirmacion", MessageBoxButtons.YesNo,
				MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;

			double dTotal = double.Parse(txtFactura.Value.ToString());
			double dFee = double.Parse(txtFee.Value.ToString());
			double dSubTotalIva = double.Parse(txtTotalConIva.Value.ToString())
				+ dFee;
			double dSubTotalExcento = double.Parse(txtTotalSinIva.Value.ToString())
				+ double.Parse(txtServiciosAdional.Value.ToString());
			double dIva = double.Parse(txtIva12.Value.ToString());
			double dDescuento = double.Parse(txtDescuento.Value.ToString());

			DateTime dtDesde = (DateTime) cmbDesde.Value;
			string stSelect = "";
			int IdCompraAntes = 0;
			int IdCompra = 0;
			if (txtIdCompra.Value != System.DBNull.Value)
				IdCompraAntes = int.Parse(txtIdCompra.Value.ToString());
			if (IdCompraAntes > 0)
			{
				stSelect = "Select Count(*) From Compra Where idCompra = " 
					+ IdCompraAntes.ToString() + " And Borrar = 0";
				int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, false);
				if (iCuenta == 0)
				{
					MessageBox.Show("Factura no existe o Borrada se creara otra", "Informacion");
					IdCompraAntes = 0;
				}
			}

			if (IdCompraAntes > 0)
			{
				if (DialogResult.No == MessageBox.Show("Factura ya Existe Desea Reemplazarla?",
					"Confirmacion", MessageBoxButtons.YesNo,
					MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;
				IdCompra = IdCompraAntes;
				stSelect = "DELETE DetCompra WHERE idCompra = " + IdCompraAntes.ToString();
				Funcion.EjecutaSQL(cdsCliente, stSelect);
				stSelect = "UPDATE Compra SET idCliente = " + cmbCliente.Value.ToString()
					+ ", Fecha = '" + dtDesde.ToString("yyyyMMdd") 
					+ "', FechaVencimiento = DateAdd(Day, 30, '" + dtDesde.ToString("yyyyMMdd") + "')"
					+ ", Notas = '" + txtPasajeros.Text + "', Total = "
					+ dTotal.ToString("0.00", us) + ", Iva = " + dIva.ToString("0.00", us)
					+ ", DiasCredito = 30, ContadoCredito = 2"
					+ ", SubtotalIva = " + dSubTotalIva.ToString("0.00", us) 
					+ ", SubtotalExcento = " + dSubTotalExcento.ToString("0.00", us)
					+ ", Usuario = '" + MenuLatinium.stUsuario + "'"
					+ ", Descuento = " + dDescuento.ToString("0.00", us)
					+ ", Desc3 = "  + dDescuento.ToString("0.00", us)
					+ " WHERE idCompra = " + IdCompraAntes.ToString();
			}
			else
			{
				string stNumero = Funcion.NumeraFactura(cdsCliente, 1);
				stSelect = "INSERT INTO Compra (idCliente "
					+ ", idTipoFactura, idComprobante, Fecha, FechaVencimiento"
					+ ", Numero, Notas, Total, Iva, DiasCredito, ContadoCredito"
					+ ", SubtotalIva, SubtotalExcento, Usuario, Otro, Descuento, Desc3)"
					+ " VALUES (" + cmbCliente.Value.ToString() + ", 1, 15, '" 
					+ dtDesde.ToString("yyyyMMdd") + "', "
					+ "DateAdd(Day, 30, '" + dtDesde.ToString("yyyyMMdd")+ "'), '" 
					+ stNumero + "', '" + txtPasajeros.Text + "', "
					+ dTotal.ToString("0.00", us) + ", " + dIva.ToString("0.00", us)
					+ ", 30, 2, " + dSubTotalIva.ToString("0.00", us) + ", " 
					+ dSubTotalExcento.ToString("0.00", us) + ", '"
					+ MenuLatinium.stUsuario + "', '" + txtNumero.Value.ToString() 
					+ "', " + dDescuento.ToString("0.00", us)
					+ ", "  + dDescuento.ToString("0.00", us)
					+ ")";
				Funcion.EjecutaSQL(cdsReservacion, stSelect);

				stSelect = "SELECT TOP 1 idCompra FROM Compra Order By IdCompra DESC";
				IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect);

				txtIdCompra.Value = IdCompra;
				stSelect = "Update Reservacion SET idFactura = " + IdCompra.ToString()
					+ " Where idReservacion = " + txtIdReservacion.Value.ToString();
				Funcion.EjecutaSQL(cdsReservacion, stSelect);
			}
			Funcion.EjecutaSQL(cdsCliente, stSelect);
		

			int iNumPasajeros = int.Parse(txtNumPasajeros.Value.ToString());
			double dValor = double.Parse(txtReservacion.Value.ToString())/ iNumPasajeros;
			double dComision = double.Parse(txtComision.Value.ToString())/ iNumPasajeros;
			double dImpuestos = double.Parse(txtImpuestos.Value.ToString())/ iNumPasajeros;
			double dServicios = double.Parse(txtServicios.Value.ToString())/ iNumPasajeros;
			double dIntermediacion = double.Parse(txtIntermediacion.Value.ToString())/ iNumPasajeros;

			// Valor del Servicio
			dValor = dValor - (dComision + (dImpuestos + dServicios) * 0 + dIntermediacion);
			if (dValor < 0) 
			{
				errorProvider.SetError(txtReservacion, "Valor de Reserva muy Bajo");
				return;
			}
			if (dValor > 0)
			{
				stSelect = "INSERT INTO DetCompra (idCompra "
					+ ", Cantidad, idArticulo, Precio, Impuesto, Notas)"
					+ " VALUES (" + IdCompra.ToString() + ", "
					+ iNumPasajeros.ToString() + ", " + int.Parse(drConstante["idArticulo1"].ToString())
					+ ", " + dValor.ToString("0.00", us) + ", 0, '"
					+ txtReserva.Text.Trim() + "')";
				Funcion.EjecutaSQL(cdsCliente, stSelect);
			}

			if (dFee > 0)
			{
				stSelect = "Select idArticulo From Articulo Where Codigo = 'FEE'";
				int idArtFee = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
				stSelect = "INSERT INTO DetCompra (idCompra "
					+ ", Cantidad, idArticulo, Precio, Impuesto, DescuentoArt)"
					+ " VALUES (" + IdCompra.ToString() + ", 1"
					+ ", " + idArtFee.ToString()
					+ ", " + dFee.ToString("0.00", us) + ", 12, " + dDescuento.ToString() + ")";
				Funcion.EjecutaSQL(cdsCliente, stSelect);
			}

			// Valor de Intermediacion
			if (dIntermediacion > 0)
			{
				stSelect = "INSERT INTO DetCompra (idCompra "
					+ ", Cantidad, idArticulo, Precio, Impuesto)"
					+ " VALUES (" + IdCompra.ToString() + ", "
					+ txtNumPasajeros.Value.ToString() + ", " + int.Parse(drConstante["idArticulo2"].ToString())
					+ ", " + dIntermediacion.ToString("0.00", us) + ", 12)";
				Funcion.EjecutaSQL(cdsCliente, stSelect);
			}

			// Valor de Impuestos extranjeros
			if (dImpuestos > 0)
			{
				stSelect = "INSERT INTO DetCompra (idCompra "
					+ ", Cantidad, idArticulo, Precio, Impuesto, Notas)"
					+ " VALUES (" + IdCompra.ToString() + ", "
					+ txtNumPasajeros.Value.ToString() + ", " + int.Parse(drConstante["idArticulo3"].ToString())
					+ ", " + dImpuestos.ToString("0.00", us) + ", 0, '"
					+ txtNotaImpuesto.Text.Trim() + "')";
				Funcion.EjecutaSQL(cdsCliente, stSelect);
			}

			// Valor de Servicios Adicionales
			if (dServicios > 0)
			{
				stSelect = "INSERT INTO DetCompra (idCompra "
					+ ", Cantidad, idArticulo, Precio, Impuesto)"
					+ " VALUES (" + IdCompra.ToString() + ", "
					+ txtNumPasajeros.Value.ToString() + ", " + int.Parse(drConstante["idArticulo4"].ToString())
					+ ", " + dServicios.ToString("0.00", us) + ", 0)";
				Funcion.EjecutaSQL(cdsCliente, stSelect);
			}

			// Graba idCompra en Reservacion
			stSelect = "Exec FacturaTotales " + IdCompra;
			Funcion.EjecutaSQL(cdsReservacion, stSelect);

			stSelect = "Update Reservacion SET idFactura = "+ IdCompra.ToString() 
				+ " Where idReservacion = " + txtIdReservacion.Value.ToString();
			Funcion.EjecutaSQL(cdsReservacion, stSelect);
		}

		private void barraDatoSQL1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btCalcular_Click()
		{
		
		}

		private void barraDatoSQL1_Load_1(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Restaurar la Reservacion?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) ) 
				return;
			drReserva["Borrar"] = false;
			cdsReservacion.Update();
			lblAnulado.Visible = false;
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] = 0"));
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (txtIdCompra.Value == System.DBNull.Value || int.Parse(txtIdCompra.Value.ToString()) == 0)
			{
				MessageBox.Show("No ha creado una factura", "Informacion");
				return;
			}
//			Compra miCompra = new Compra(1, int.Parse(txtIdCompra.Value.ToString()));
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void btReserva_Click(object sender, System.EventArgs e)
		{
			if (FormReserva.Text.Trim().Length == 0) return;
			string stSelect = "Select " +  FormReserva.Text;
			double dValor = Funcion.diEscalarSQL(cdsReservacion, stSelect);
			txtReservacion.Value = decimal.Parse(dValor.ToString("0.00", us));
			txtReservacion.Select();
			txtReservacion.Refresh();
		}

		private void btCalcImpuesto_Click(object sender, System.EventArgs e)
		{
			if (FormImpuesto.Text.Trim().Length == 0) return;
			string stSelect = "Select " +  FormImpuesto.Text;
			double dValor = Funcion.diEscalarSQL(cdsReservacion, stSelect);
			txtImpuestos.Value = decimal.Parse(dValor.ToString("0.00", us));
			txtImpuestos.Select();
			txtImpuestos.Refresh();
		}

		private void btSubencion_Click(object sender, System.EventArgs e)
		{
			if (FormSubencion.Text.Trim().Length == 0) return;
			string stSelect = "Select " +  FormSubencion.Text;
			double dValor = Funcion.diEscalarSQL(cdsReservacion, stSelect);
			txtSubcidio.Value = decimal.Parse(dValor.ToString("0.00", us));
			txtSubcidio.Select();
			txtSubcidio.Refresh();		
		}

		private void FormSubencion_Validated(object sender, System.EventArgs e)
		{
			btSubencion_Click(this, e);
		}

		private void FormImpuesto_Validated(object sender, System.EventArgs e)
		{
			btCalcImpuesto_Click(this, e);
		}

		private void FormReserva_Validated(object sender, System.EventArgs e)
		{
			btReserva_Click(this, e);
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Reservacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Reservacion'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void rbContado_CheckedChanged(object sender, System.EventArgs e)
		{
		}

		private void rbCredito_CheckedChanged(object sender, System.EventArgs e)
		{
		}
		#region calculo fee automatico
		private void btCalculoFee_Click(object sender, System.EventArgs e)
		{
			int i = CaluloFee(); 
		}
		private int CaluloFee()
		{
				
			int retorna = 0;					
			double ValorAntesIva = Fee(double.Parse( txtReservacion.Value.ToString()));								
			txtFee.Value = ValorAntesIva;
			txtDescuento.Value = ValorAntesIva;			
			
			if ((double) txtCredito1.Value == (double) txtReservacion.Value )			
			{
				txtDescuento.Value = 0;			
			}
			else
			{
				if ((double) txtCredito1.Value  > 0)
				{
					ValorAntesIva = FeeCred(double.Parse( txtCredito1.Value.ToString()));		
					txtDescuento.Value = double.Parse(txtFee.Value.ToString()) - ValorAntesIva;			

				}
			}

			if(double.Parse(txtDescuento.Value.ToString()) < 0 ) txtDescuento.Value = 0;
			
										
			return retorna;
		}
		private double FeeCred(double valor)
		{	
			double PorcFee =0;	
			if (Funcion.bEjecutaSQL(cdsCliente, "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'AGPFEE'"))
			{
				PorcFee = Funcion.dEscalarSQL(cdsCliente,"Select IsNull(Valor, 0) From SeteoG Where Codigo = 'AGPFEE'");
			}
			double ValorAntesIva =0;		
			double Ingreso = valor;
			double calduloFijo = (Ingreso*(2.40 / 100));
			double SumaICF = Ingreso+calduloFijo;
			double dIVA = Funcion.dEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
			ValorAntesIva = (SumaICF * (PorcFee/100))/dIVA;
			return ValorAntesIva;

		}
		private double Fee(double valor)
		{	
				double PorcFee =0;	
			if (Funcion.bEjecutaSQL(cdsCliente, "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'AGPFEE'"))
			{
				PorcFee = Funcion.dEscalarSQL(cdsCliente,"Select IsNull(Valor, 0) From SeteoG Where Codigo = 'AGPFEE'");
			}
			double ValorAntesIva =0;		
			double Ingreso = valor;
			double calduloFijo = (Ingreso*(2.40 / 100))+ double.Parse(txtImpuestos.Value.ToString());
			double SumaICF = Ingreso+calduloFijo;
			double dIVA = Funcion.dEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
			ValorAntesIva = (SumaICF * (PorcFee/100))/dIVA;
			return ValorAntesIva;

		}

		#endregion calculo fee automatico


		#region Fee Tarjeta de Credito
		private void btGeneraComprobante_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			CalculoFeeTarjeCred();
			Cursor = Cursors.Default;
		}
		private void  CalculoFeeTarjeCred()
		{
			if (txtCredito2.Value== DBNull.Value || txtContado2.Value== DBNull.Value) return;
		
			double PorcFeeTC=0, FeeTC=0, vCredCon =0, Fee =0;
			
			if (Funcion.bEjecutaSQL(cdsCliente, "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'AGPTCFEE'"))
			{
				PorcFeeTC = Funcion.dEscalarSQL(cdsCliente,"Select IsNull(Valor, 0) From SeteoG Where Codigo = 'AGPTCFEE'");
			}
			vCredCon = double.Parse( txtCredito2.Value.ToString());
			
			

			if((double) txtCredito1.Value < (double) txtCredito2.Value )
			{
				// 1 calculo cuando el valor de Fee correspone a efectivo
				//   Pero existe pago tambien con tarjeta de credito
				//	Emitir Factura adicional
				//vCredCon = vCredCon -  double.Parse( txtCredito1.Value.ToString());
				double dIVA = Funcion.dEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
				FeeTC = (vCredCon -(vCredCon/PorcFeeTC))/dIVA;
				FeeTC = FeeTC - ((double)txtFee.Value - (double)txtDescuento.Value);

				if (DialogResult.No == MessageBox.Show("Desea Generar Factura Adicional?",
					"Confirmacion", MessageBoxButtons.YesNo,
					MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;
				nuevaFactNotaCredito(1,FeeTC,15);


			}
			if((double) txtCredito1.Value > (double) txtCredito2.Value )					
			{
				// 2 calculo cuando el valor de Fee correspone a tarjeta de credito
				//   Pero existe pago tambien con efectivo
				//	Emitir una Nota de Credito
				double dIVA = Funcion.dEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
			
					FeeTC = (vCredCon -(vCredCon/PorcFeeTC))/dIVA;
				  Fee =  double.Parse( txtFee.Value.ToString());
					FeeTC = ((double)txtFee.Value - (double)txtDescuento.Value) - FeeTC;
				if (DialogResult.No == MessageBox.Show("Desea Generar Nota de Credito?",
					"Confirmacion", MessageBoxButtons.YesNo,
					MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;

				nuevaFactNotaCredito(5,FeeTC,4);
			}		
			if ((double) txtCredito1.Value == (double) txtCredito2.Value)
			{
				MessageBox.Show("No Genera...");

			}
		}
		private void nuevaFactNotaCredito(int IdTipoFact, double valor, int idComprobante)
			{
			string stSelect ="";
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			int idCompraAuxAnt=0, IdCompraAux = 0;
				if (txtIdCompraAux.Value != System.DBNull.Value)
					idCompraAuxAnt = int.Parse(txtIdCompraAux.Value.ToString());
				if (idCompraAuxAnt > 0)
				{
					stSelect = "Select Count(*) From Compra Where idCompra = " 
						+ idCompraAuxAnt.ToString() + " And Borrar = 0";
					int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, false);
					if (iCuenta == 0)
					{
						MessageBox.Show("Factura no existe o Borrada se creara otra", "Informacion");
						idCompraAuxAnt = 0;
					}
				}
				if (idCompraAuxAnt > 0)
				{
					IdCompraAux = idCompraAuxAnt;
					if (DialogResult.No == MessageBox.Show("Factura ya Existe Desea Reemplazarla?",
						"Confirmacion", MessageBoxButtons.YesNo,
						MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;
					stSelect = "DELETE DetCompra WHERE idCompra = " + idCompraAuxAnt.ToString();
					Funcion.EjecutaSQL(cdsCliente, stSelect);
					stSelect = "UPDATE Compra SET idCliente = " + cmbCliente.Value.ToString()
						+ ", Fecha = '" + dtDesde.ToString("yyyyMMdd") 
						+ "', FechaVencimiento = DateAdd(Day, 30, '" + dtDesde.ToString("yyyyMMdd") + "')"
						+ ", Notas = '" + txtPasajeros.Text + "' "
						+ ", DiasCredito = 30, ContadoCredito = 2"
						+ ", Usuario = '" + MenuLatinium.stUsuario + "'"
						+ ", idTipoFactura = " + IdTipoFact 
						+ ", idComprobante = " + idComprobante 
						+ " WHERE idCompra = " + idCompraAuxAnt.ToString();

					Funcion.EjecutaSQL(cdsReservacion, stSelect);
				}
				else
				{
					string stNumero = Funcion.NumeraFactura(cdsCliente, IdTipoFact);
					stSelect = "INSERT INTO Compra (idCliente "
						+ ", idTipoFactura, idComprobante, Fecha, FechaVencimiento"
						+ ", Numero, Notas,DiasCredito, ContadoCredito"
						+ ", Usuario, Otro)"
						+ " VALUES (" + cmbCliente.Value.ToString() + ", "+ IdTipoFact+","+ idComprobante +", '" 
						+ dtDesde.ToString("yyyyMMdd") + "', "
						+ "DateAdd(Day, 30, '" + dtDesde.ToString("yyyyMMdd")+ "'), '" 
						+ stNumero + "', '" + txtPasajeros.Text + "' "						
						+ ", 30, 2, '"
						+ MenuLatinium.stUsuario + "', '" + txtNumero.Value.ToString() 
						+ "')";

					Funcion.EjecutaSQL(cdsReservacion, stSelect);

					stSelect = "SELECT TOP 1 idCompra FROM Compra Order By IdCompra DESC";
					IdCompraAux = Funcion.iEscalarSQL(cdsCliente, stSelect);

					txtIdCompraAux.Value = IdCompraAux;
					stSelect = "Update Reservacion SET idFacturaAux = " + IdCompraAux.ToString()
						+ " Where idReservacion = " + txtIdReservacion.Value.ToString();
					Funcion.EjecutaSQL(cdsReservacion, stSelect);
				}
			
				// Valor FEE			
				if (valor > 0)
				{
					stSelect = "Select idArticulo From Articulo Where Codigo = 'FEE'";
					int idArtFee = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
					stSelect = "INSERT INTO DetCompra (idCompra "
						+ ", Cantidad, idArticulo, Precio, Impuesto)"
						+ " VALUES (" + IdCompraAux.ToString() + ", 1"
						+ ", " + idArtFee.ToString()
						+ ", " + valor.ToString("0.00", us) + ", 12 )";
					Funcion.EjecutaSQL(cdsCliente, stSelect);
				}

			Funcion.EjecutaSQL(cdsCliente,"Exec FacturaTotales " + IdCompraAux);
			}


		

		private void btGeneraComprobante_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right )return;

			if (txtIdCompraAux.Value == System.DBNull.Value || int.Parse(txtIdCompraAux.Value.ToString()) == 0)
			{
				MessageBox.Show("No ha creado una factura", "Informacion");
				return;
			}
			int tipoF = 1;

			if((double) txtCredito1.Value > (double) txtCredito2.Value )	tipoF = 5;	
//			Compra miCompra = new Compra(tipoF, int.Parse(txtIdCompraAux.Value.ToString()));
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		#endregion Fee Tarjeta de Credito

		private void Copiar_Click(object sender, System.EventArgs e)
		{
			if (txtReservacion.Value != System.DBNull.Value && txtImpuestos.Value != System.DBNull.Value  )
			{
				double res = double.Parse(txtReservacion.Value.ToString());
				double imp = double.Parse(txtImpuestos.Value.ToString());
				txtCredito1.Value = (res + imp );
			}
		}
	}
}
