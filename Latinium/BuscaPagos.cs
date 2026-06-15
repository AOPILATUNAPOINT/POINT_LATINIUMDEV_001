using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaPagos : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCompra;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtVencimiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtEntrega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtros;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbNombre;
		private C1.Data.C1DataView cdvCliente;
		private System.ComponentModel.IContainer components=null;

		public BuscaPagos()
		{
			InitializeComponent();
		}
		int IdSigno = 1;
		public BuscaPagos(int idTipoF)
		{
			InitializeComponent();
			IdSigno = idTipoF;
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLot");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoPer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMoneda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDolar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Refrendo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagosCada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prioridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FletePorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BienServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioInternac");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaRetDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprob1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depart");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bulto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloquear");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CifIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Propina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutoImpSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNiffIngProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NiifValorIng");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IntImplicito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TasaCambio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoDif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("subTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAux");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipoObra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nivePrioridadRuta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaMonitoreado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sistemaSeguridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoNIIF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactAfectaCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraCosto");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			this.cdvCompra = new C1.Data.C1DataView();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtVencimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtOtros = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbNombre = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCliente = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVencimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.CaseSensitive = false;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvCompra, "idCompra"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance1;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Que Empiece";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Exacta";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem1);
			this.rdgTipoBusqueda.Items.Add(valueListItem2);
			this.rdgTipoBusqueda.Items.Add(valueListItem3);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(288, 60);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 42;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// txtCodigo
			// 
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 43;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(136, 69);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(128, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 45;
			this.txtFecha.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtVencimiento
			// 
			this.txtVencimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.txtVencimiento.DateButtons.Add(dateButton2);
			this.txtVencimiento.Format = "dd/MMM/yyyy";
			this.txtVencimiento.Location = new System.Drawing.Point(136, 95);
			this.txtVencimiento.Name = "txtVencimiento";
			this.txtVencimiento.NonAutoSizeHeight = 23;
			this.txtVencimiento.Size = new System.Drawing.Size(128, 21);
			this.txtVencimiento.SpinButtonsVisible = true;
			this.txtVencimiento.TabIndex = 46;
			this.txtVencimiento.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtEntrega
			// 
			this.txtEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.txtEntrega.DateButtons.Add(dateButton3);
			this.txtEntrega.Format = "dd/MMM/yyyy";
			this.txtEntrega.Location = new System.Drawing.Point(136, 121);
			this.txtEntrega.Name = "txtEntrega";
			this.txtEntrega.NonAutoSizeHeight = 23;
			this.txtEntrega.Size = new System.Drawing.Size(128, 21);
			this.txtEntrega.SpinButtonsVisible = true;
			this.txtEntrega.TabIndex = 47;
			this.txtEntrega.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtComprobante
			// 
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(136, 146);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(128, 22);
			this.txtComprobante.TabIndex = 48;
			// 
			// txtOtros
			// 
			this.txtOtros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtros.Location = new System.Drawing.Point(136, 172);
			this.txtOtros.Name = "txtOtros";
			this.txtOtros.Size = new System.Drawing.Size(128, 22);
			this.txtOtros.TabIndex = 49;
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance2;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Nombre";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Código";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Fecha";
			valueListItem7.DataValue = 4;
			valueListItem7.DisplayText = "Vencimiento";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "Entrega";
			valueListItem9.DataValue = 6;
			valueListItem9.DisplayText = "Comprobante";
			valueListItem10.DataValue = 7;
			valueListItem10.DisplayText = "Otros";
			valueListItem11.DataValue = 8;
			valueListItem11.DisplayText = "Total";
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.Items.Add(valueListItem6);
			this.rdgOpcion.Items.Add(valueListItem7);
			this.rdgOpcion.Items.Add(valueListItem8);
			this.rdgOpcion.Items.Add(valueListItem9);
			this.rdgOpcion.Items.Add(valueListItem10);
			this.rdgOpcion.Items.Add(valueListItem11);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 207);
			this.rdgOpcion.TabIndex = 51;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvCompra;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.AddButtonCaption = "Cliente";
			ultraGridColumn1.Header.VisiblePosition = 43;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 95;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 84;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 87;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Cliente";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn6.Header.VisiblePosition = 40;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 86;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Centro Costo";
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn9.Header.Caption = "Proyecto";
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Header.VisiblePosition = 76;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 47;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 46;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 48;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 41;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 33;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 49;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 44;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 45;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 82;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 78;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Format = "dd/MMM/yyyy";
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn22.Header.VisiblePosition = 27;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Format = "dd/MMM/yyyy";
			ultraGridColumn23.Header.Caption = "Vencimiento";
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn24.Header.VisiblePosition = 68;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.Caption = "Número";
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn26.Header.VisiblePosition = 36;
			ultraGridColumn26.Hidden = true;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance4;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn28.Header.VisiblePosition = 23;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 24;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 25;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 26;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 30;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 75;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 52;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 51;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 22;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 17;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 19;
			ultraGridColumn38.Hidden = true;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance5;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 3;
			ultraGridColumn40.Header.VisiblePosition = 59;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 57;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 18;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 58;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 69;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 67;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 70;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 61;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 71;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 34;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 74;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 73;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Format = "dd/MMM/yyyy";
			ultraGridColumn52.Header.Caption = "Entrega";
			ultraGridColumn52.Header.VisiblePosition = 6;
			ultraGridColumn53.Header.VisiblePosition = 8;
			ultraGridColumn54.Header.VisiblePosition = 66;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 65;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 64;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 21;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 39;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 63;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 7;
			ultraGridColumn61.Header.VisiblePosition = 15;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 37;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 60;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 31;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 12;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 11;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 42;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 29;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 28;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 20;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 62;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 14;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 13;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 53;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 54;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 55;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 35;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 38;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 50;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 56;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 32;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 16;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 72;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 77;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 79;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 80;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 81;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 83;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 85;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 88;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 89;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 90;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 94;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 91;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 92;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 96;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 97;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 98;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 99;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 100;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 101;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 102;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 103;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 104;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 105;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 106;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 107;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 108;
			ultraGridColumn110.Header.VisiblePosition = 109;
			ultraGridColumn111.Header.VisiblePosition = 110;
			ultraGridColumn112.Header.VisiblePosition = 111;
			ultraGridColumn113.Header.VisiblePosition = 112;
			ultraGridColumn114.Header.VisiblePosition = 113;
			ultraGridColumn115.Header.VisiblePosition = 114;
			ultraGridColumn116.Header.VisiblePosition = 115;
			ultraGridColumn117.Header.VisiblePosition = 116;
			ultraGridColumn118.Header.VisiblePosition = 117;
			ultraGridColumn119.Header.VisiblePosition = 118;
			ultraGridColumn120.Header.VisiblePosition = 119;
			ultraGridColumn121.Header.VisiblePosition = 120;
			ultraGridColumn122.Header.VisiblePosition = 121;
			ultraGridColumn123.Header.VisiblePosition = 122;
			ultraGridColumn124.Header.VisiblePosition = 123;
			ultraGridColumn125.Header.VisiblePosition = 124;
			ultraGridColumn126.Header.VisiblePosition = 125;
			ultraGridColumn127.Header.VisiblePosition = 126;
			ultraGridColumn128.Header.VisiblePosition = 127;
			ultraGridColumn129.Header.VisiblePosition = 128;
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
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 233);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(504, 198);
			this.ultraGrid1.TabIndex = 52;
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DisplayMember = "Nombre";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(48, 284);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(280, 86);
			this.ultraDropDown1.TabIndex = 53;
			this.ultraDropDown1.Text = "ultraDropDown1";
			this.ultraDropDown1.ValueMember = "idCliente";
			this.ultraDropDown1.Visible = false;
			// 
			// txtTotal
			// 
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(136, 198);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(128, 22);
			this.txtTotal.TabIndex = 54;
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(440, 129);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 57;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(440, 95);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 56;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(440, 60);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 55;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btExcel
			// 
			this.btExcel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btExcel.Location = new System.Drawing.Point(440, 172);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(75, 25);
			this.btExcel.TabIndex = 58;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			ultraGridBand2.AddButtonCaption = "Cliente";
			ultraGridColumn130.Header.VisiblePosition = 1;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 2;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn132.Width = 249;
			ultraGridColumn133.Header.VisiblePosition = 3;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 4;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 5;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 6;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 7;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 8;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 9;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(184, 319);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(168, 69);
			this.cmbProyecto.TabIndex = 59;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.CaseSensitive = false;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cmbSubProyecto
			// 
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataMember = "SubProyecto";
			this.cmbSubProyecto.DataSource = this.cdsProyecto;
			ultraGridBand3.AddButtonCaption = "Cliente";
			ultraGridColumn140.Header.VisiblePosition = 1;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 2;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 0;
			ultraGridColumn142.Width = 249;
			ultraGridColumn143.Header.VisiblePosition = 3;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 4;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 5;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 6;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 7;
			ultraGridColumn147.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto.Location = new System.Drawing.Point(360, 327);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(168, 69);
			this.cmbSubProyecto.TabIndex = 60;
			this.cmbSubProyecto.Text = "cmbSubProyecto";
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.Visible = false;
			// 
			// cmbNombre
			// 
			this.cmbNombre.Location = new System.Drawing.Point(136, 16);
			this.cmbNombre.Name = "cmbNombre";
			// 
			// cmbNombre.Properties
			// 
			this.cmbNombre.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbNombre.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbNombre.Properties.DisplayMember = "Nombre";
			this.cmbNombre.Properties.NullText = "";
			this.cmbNombre.Properties.PopupWidth = 300;
			this.cmbNombre.Properties.ShowFooter = false;
			this.cmbNombre.Properties.ShowHeader = false;
			this.cmbNombre.Properties.ValueMember = "idCliente";
			this.cmbNombre.Size = new System.Drawing.Size(288, 20);
			this.cmbNombre.TabIndex = 65;
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// BuscaPagos
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 445);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtOtros);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtEntrega);
			this.Controls.Add(this.txtVencimiento);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaPagos";
			this.Text = "Búsqueda de Pagos";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVencimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			idBusca.Width =0;
			txtFecha.Value = DateTime.Today;
			txtEntrega.Value = DateTime.Today;
			txtVencimiento.Value = DateTime.Today;
			rdgOpcion.Value = 1;
			int idTipo = 1;
			if (IdSigno != 1) idTipo = 4;
			string stExec = "Exec ClienteCargaNombre " + idTipo;
			cmbNombre.Properties.DataSource = Funcion.dvProcedimiento(cdsCompra, stExec);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";

			strFiltro = "";
			DateTime dtFecha;
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (this.cmbNombre.EditValue == null)
					{
						this.errorProvider.SetError(this.cmbNombre, "Ingreses Nombre");
						return;
					}
					this.errorProvider.SetError(this.cmbNombre, "");

					strFiltro = "idCliente = " + this.cmbNombre.EditValue.ToString().Trim();
					break;
				case 1:
					if (this.txtCodigo.Value == null)
					{
						this.errorProvider.SetError(this.txtCodigo, "Ingreses Codigo");
						return;
					}
					this.errorProvider.SetError(this.txtCodigo, "");
					stBusca = this.txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Numero LIKE '" + strAntes + stBusca + strDespues;

					break;
				case 2:
					break;
				case 3:
					dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = "Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 4:
					dtFecha = (DateTime) this.txtEntrega.Value;
					strFiltro = "FechaVencimiento >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and FechaVencimiento < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 5:
					dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = "FechaEntrega >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and FechaEntrega < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 6:
					if (this.txtComprobante.Value == null)
					{
						this.errorProvider.SetError(this.txtComprobante, "Ingreses Comprobante");
						return;
					}
					this.errorProvider.SetError(this.txtComprobante, "");
					stBusca = this.txtComprobante.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Comprobante LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 7:
					if (this.txtOtros.Value == null)
					{
						this.errorProvider.SetError(this.txtOtros, "Ingreses Otro");
						return;
					}
					this.errorProvider.SetError(this.txtOtros, "");

					stBusca = this.txtOtros.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Otro LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 8:
					strFiltro = "Total >=" + this.txtTotal.Value.ToString().Trim();
					break;
			}
//			MessageBox.Show(strFiltro);
			Cursor = Cursors.WaitCursor;
			string stFiltro = strFiltro + " And Compra.Borrar = 0 ";
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			{
				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
				fcTotal.Add(fcHija);
			}

			this.cdsCompra.Fill(fcTotal, false);

			if (this.cdvCompra.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				Cursor = Cursors.Default;
				return;
			}

			if (this.cdvCompra.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
			Cursor = Cursors.Default;
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			C1.Data.SchemaObjects.DataSetDef dataSetDef = 
				e.DataSet.Schema.DataSetDefs["dsCompra"];

			string stFiltro = "";
			if (IdSigno == 1)
				stFiltro = "([idTipoFactura] = 1 Or [idTipoFactura] = 12 Or [idTipoFactura] = 99) And [ContadoCredito] = 2";
			else
				stFiltro = "([idTipoFactura] = 4 Or [idTipoFactura] = 11) And [ContadoCredito] = 2";

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Compra"],	stFiltro));
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.cmbNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
//			this.txtNumero.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtVencimiento.Enabled = false;
			this.txtEntrega.Enabled = false;
			this.txtComprobante.Enabled = false;
			this.txtOtros.Enabled = false;
			this.txtTotal.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.cmbNombre.Enabled = true;
					this.cmbNombre.Select();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					break;
				case 2:
//					this.txtNumero.Enabled = true;
//					this.txtNumero.Select();
					break;
				case 3:
					this.txtFecha.Enabled = true;
					this.txtFecha.Select();
					break;
				case 4:
					this.txtVencimiento.Enabled = true;
					this.txtVencimiento.Select();
					break;
				case 5:
					this.txtEntrega.Enabled = true;
					this.txtEntrega.Select();
					break;
				case 6:
					this.txtComprobante.Enabled = true;
					this.txtComprobante.Select();
					break;
				case 7:
					this.txtOtros.Enabled = true;
					this.txtOtros.Select();
					break;
				case 8:
					this.txtTotal.Enabled = true;
					this.txtTotal.Select();
					break;
			}

		}
	}
}
