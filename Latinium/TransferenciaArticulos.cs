using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TransferenciaArticulos.
	/// </summary>
	public class TransferenciaArticulos : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTmpArt;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodigo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRecibe;
		private Infragistics.Win.Misc.UltraButton btTransformar;
		private Infragistics.Win.Misc.UltraButton btQuitarTodo;
		private Infragistics.Win.Misc.UltraButton btAgregarTodo;
		private Infragistics.Win.Misc.UltraButton btQuitar;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btCambiaER;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.Misc.UltraButton btTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActualizaPrecio;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkServicio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label lblProyecto;
		private System.Windows.Forms.Label lblCaja;
		private Infragistics.Win.Misc.UltraButton btTSP;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TransferenciaArticulos()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLot");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoPer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_AnticipoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_MultasT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FondoGT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngAfectaCostoNiif");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FaltaRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_FaltaOriginal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IRBP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_IdOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Procesado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "cmbArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0, "cmbCodigo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_idEducativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaDscto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturable");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_idEducativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaDscto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturable");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn225 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn226 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn227 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn228 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn229 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn230 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn231 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn232 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn233 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn234 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn235 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn236 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn237 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn238 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn239 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn240 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn241 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn242 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn243 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn244 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn245 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn246 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn247 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn248 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn249 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn250 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn251 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn252 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn253 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn254 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn255 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn256 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn257 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn258 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn259 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn260 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn261 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn262 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn263 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn264 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn265 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn266 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn267 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn268 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn269 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn270 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn271 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn272 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn273 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn274 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn275 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn276 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn277 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn278 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn279 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn280 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn281 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn282 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn283 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn284 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn285 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn286 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn287 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn288 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn289 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn290 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn291 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn292 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn293 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbTipoRecibe = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsTmpArt = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkActualizaPrecio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.btQuitarTodo = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btTransformar = new Infragistics.Win.Misc.UltraButton();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAgregarTodo = new Infragistics.Win.Misc.UltraButton();
			this.btQuitar = new Infragistics.Win.Misc.UltraButton();
			this.btCambiaER = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.btTotal = new Infragistics.Win.Misc.UltraButton();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.label5 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkServicio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblProyecto = new System.Windows.Forms.Label();
			this.lblCaja = new System.Windows.Forms.Label();
			this.btTSP = new Infragistics.Win.Misc.UltraButton();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTmpArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			this.SuspendLayout();
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
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Cliente";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Compra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 36;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 95;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 84;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 87;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 35;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 33;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 86;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 41;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 42;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 76;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 40;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 39;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 43;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 34;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 45;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 44;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 37;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 38;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 82;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 78;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Format = "dd/MMM/yyyy";
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 82;
			ultraGridColumn22.Header.VisiblePosition = 26;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 29;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 68;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn26.Header.VisiblePosition = 53;
			ultraGridColumn26.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance2;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn28.Header.VisiblePosition = 16;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 17;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 18;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 19;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 22;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 75;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 48;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 47;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 23;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 11;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 13;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 63;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 57;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 55;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 12;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 56;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 69;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 67;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 70;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 59;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 71;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 28;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 74;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 73;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 27;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 4;
			ultraGridColumn54.Header.VisiblePosition = 66;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 65;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 62;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 15;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 32;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 61;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 3;
			ultraGridColumn61.Header.VisiblePosition = 9;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 30;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 58;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 24;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 6;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 5;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 64;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 21;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 20;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 14;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 60;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 7;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 49;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 50;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 51;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 52;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 31;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 46;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 54;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 25;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 10;
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
			ultraGridColumn130.Header.VisiblePosition = 129;
			ultraGridColumn131.Header.VisiblePosition = 130;
			ultraGridColumn132.Header.VisiblePosition = 131;
			ultraGridColumn133.Header.VisiblePosition = 132;
			ultraGridColumn134.Header.VisiblePosition = 133;
			ultraGridColumn135.Header.VisiblePosition = 134;
			ultraGridColumn136.Header.VisiblePosition = 135;
			ultraGridColumn137.Header.VisiblePosition = 136;
			ultraGridColumn138.Header.VisiblePosition = 137;
			ultraGridColumn139.Header.VisiblePosition = 138;
			ultraGridColumn140.Header.VisiblePosition = 139;
			ultraGridColumn141.Header.VisiblePosition = 140;
			ultraGridColumn142.Header.VisiblePosition = 141;
			ultraGridColumn143.Header.VisiblePosition = 142;
			ultraGridColumn144.Header.VisiblePosition = 143;
			ultraGridColumn145.Header.VisiblePosition = 144;
			ultraGridColumn146.Header.VisiblePosition = 145;
			ultraGridColumn147.Header.VisiblePosition = 146;
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
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.ExtendedAutoDrag;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(288, 192);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseUp);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
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
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			this.cmbTipo.DisplayMember = "Nombre";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(64, 9);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(176, 21);
			this.cmbTipo.TabIndex = 3;
			this.cmbTipo.ValueMember = "";
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(248, 9);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(48, 24);
			this.btFiltro.TabIndex = 4;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			this.btFiltro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btFiltro_MouseUp);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(26, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Tipo";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(360, 9);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(26, 16);
			this.label3.TabIndex = 10;
			this.label3.Text = "Tipo";
			// 
			// cmbTipoRecibe
			// 
			this.cmbTipoRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRecibe.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRecibe.DisplayLayout.AutoFitColumns = true;
			this.cmbTipoRecibe.DisplayMember = "Nombre";
			this.cmbTipoRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRecibe.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRecibe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRecibe.Location = new System.Drawing.Point(408, 9);
			this.cmbTipoRecibe.Name = "cmbTipoRecibe";
			this.cmbTipoRecibe.Size = new System.Drawing.Size(176, 21);
			this.cmbTipoRecibe.TabIndex = 8;
			this.cmbTipoRecibe.ValueMember = "";
			this.cmbTipoRecibe.TextChanged += new System.EventHandler(this.cmbTipoRecibe_TextChanged);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "Band 0";
			this.ultraGrid2.DataSource = this.udsTmpArt;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance7;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn148.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn148.Header.Caption = "Articulo";
			ultraGridColumn148.Header.VisiblePosition = 0;
			ultraGridColumn148.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn148.Width = 52;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn149.CellAppearance = appearance8;
			ultraGridColumn149.Format = "#,##0.00";
			ultraGridColumn149.Header.VisiblePosition = 2;
			ultraGridColumn149.Width = 28;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn150.CellAppearance = appearance9;
			ultraGridColumn150.Format = "#,##0.00";
			ultraGridColumn150.Header.VisiblePosition = 3;
			ultraGridColumn150.Width = 30;
			ultraGridColumn151.Header.VisiblePosition = 4;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn151.Width = 66;
			ultraGridColumn152.Header.VisiblePosition = 5;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn152.Width = 35;
			ultraGridColumn153.Header.VisiblePosition = 6;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn153.Width = 49;
			ultraGridColumn154.Header.VisiblePosition = 7;
			ultraGridColumn154.Width = 52;
			ultraGridColumn155.Header.VisiblePosition = 8;
			ultraGridColumn155.Width = 63;
			ultraGridColumn156.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn156.DataType = typeof(int);
			ultraGridColumn156.Header.Caption = "Código";
			ultraGridColumn156.Header.VisiblePosition = 1;
			ultraGridColumn156.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn156.Width = 42;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
																										 ultraGridColumn155,
																										 ultraGridColumn156});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(352, 104);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(288, 200);
			this.ultraGrid2.TabIndex = 11;
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			// 
			// udsTmpArt
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn7.DefaultValue = 0;
			this.udsTmpArt.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2,
																													ultraDataColumn3,
																													ultraDataColumn4,
																													ultraDataColumn5,
																													ultraDataColumn6,
																													ultraDataColumn7,
																													ultraDataColumn8});
			// 
			// chkActualizaPrecio
			// 
			this.chkActualizaPrecio.Location = new System.Drawing.Point(368, 72);
			this.chkActualizaPrecio.Name = "chkActualizaPrecio";
			this.chkActualizaPrecio.Size = new System.Drawing.Size(120, 22);
			this.chkActualizaPrecio.TabIndex = 12;
			this.chkActualizaPrecio.Text = "Actualizar Precios";
			// 
			// btAgregar
			// 
			this.btAgregar.Location = new System.Drawing.Point(312, 155);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(32, 25);
			this.btAgregar.TabIndex = 13;
			this.btAgregar.Text = "->";
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// btQuitarTodo
			// 
			this.btQuitarTodo.Location = new System.Drawing.Point(312, 224);
			this.btQuitarTodo.Name = "btQuitarTodo";
			this.btQuitarTodo.Size = new System.Drawing.Size(32, 25);
			this.btQuitarTodo.TabIndex = 14;
			this.btQuitarTodo.Text = "<<";
			this.btQuitarTodo.Click += new System.EventHandler(this.btQuitarTodo_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btTransformar
			// 
			this.btTransformar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btTransformar.Location = new System.Drawing.Point(360, 312);
			this.btTransformar.Name = "btTransformar";
			this.btTransformar.Size = new System.Drawing.Size(88, 25);
			this.btTransformar.TabIndex = 15;
			this.btTransformar.Text = "Transformar";
			this.btTransformar.Click += new System.EventHandler(this.btTransformar_Click);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataMember = "Articulo";
			this.cmbArticulo.DataSource = this.cdsArticulo;
			ultraGridColumn157.Header.VisiblePosition = 2;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 3;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 4;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 5;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 6;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 0;
			ultraGridColumn162.Width = 176;
			ultraGridColumn163.Header.VisiblePosition = 1;
			ultraGridColumn164.Header.VisiblePosition = 7;
			ultraGridColumn165.Header.VisiblePosition = 8;
			ultraGridColumn166.Header.VisiblePosition = 9;
			ultraGridColumn167.Header.VisiblePosition = 10;
			ultraGridColumn168.Header.VisiblePosition = 11;
			ultraGridColumn169.Header.VisiblePosition = 12;
			ultraGridColumn170.Header.VisiblePosition = 13;
			ultraGridColumn171.Header.VisiblePosition = 14;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn157,
																										 ultraGridColumn158,
																										 ultraGridColumn159,
																										 ultraGridColumn160,
																										 ultraGridColumn161,
																										 ultraGridColumn162,
																										 ultraGridColumn163,
																										 ultraGridColumn164,
																										 ultraGridColumn165,
																										 ultraGridColumn166,
																										 ultraGridColumn167,
																										 ultraGridColumn168,
																										 ultraGridColumn169,
																										 ultraGridColumn170,
																										 ultraGridColumn171});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(360, 138);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(256, 86);
			this.cmbArticulo.TabIndex = 16;
			this.cmbArticulo.ValueMember = "idArticulo";
			this.cmbArticulo.Visible = false;
			// 
			// cmbCodigo
			// 
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodigo.DataMember = "Articulo";
			this.cmbCodigo.DataSource = this.cdsArticulo;
			ultraGridColumn172.Header.VisiblePosition = 2;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 3;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 4;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 5;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 6;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 1;
			ultraGridColumn177.Width = 147;
			ultraGridColumn178.Header.VisiblePosition = 0;
			ultraGridColumn179.Header.VisiblePosition = 7;
			ultraGridColumn180.Header.VisiblePosition = 8;
			ultraGridColumn181.Header.VisiblePosition = 9;
			ultraGridColumn182.Header.VisiblePosition = 10;
			ultraGridColumn183.Header.VisiblePosition = 11;
			ultraGridColumn184.Header.VisiblePosition = 12;
			ultraGridColumn185.Header.VisiblePosition = 13;
			ultraGridColumn186.Header.VisiblePosition = 14;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn172,
																										 ultraGridColumn173,
																										 ultraGridColumn174,
																										 ultraGridColumn175,
																										 ultraGridColumn176,
																										 ultraGridColumn177,
																										 ultraGridColumn178,
																										 ultraGridColumn179,
																										 ultraGridColumn180,
																										 ultraGridColumn181,
																										 ultraGridColumn182,
																										 ultraGridColumn183,
																										 ultraGridColumn184,
																										 ultraGridColumn185,
																										 ultraGridColumn186});
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCodigo.DisplayMember = "Codigo";
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(392, 164);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.Size = new System.Drawing.Size(256, 86);
			this.cmbCodigo.TabIndex = 17;
			this.cmbCodigo.ValueMember = "idArticulo";
			this.cmbCodigo.Visible = false;
			// 
			// btAgregarTodo
			// 
			this.btAgregarTodo.Location = new System.Drawing.Point(312, 121);
			this.btAgregarTodo.Name = "btAgregarTodo";
			this.btAgregarTodo.Size = new System.Drawing.Size(32, 24);
			this.btAgregarTodo.TabIndex = 18;
			this.btAgregarTodo.Text = ">>";
			this.btAgregarTodo.Click += new System.EventHandler(this.btAgregarTodo_Click);
			// 
			// btQuitar
			// 
			this.btQuitar.Location = new System.Drawing.Point(312, 190);
			this.btQuitar.Name = "btQuitar";
			this.btQuitar.Size = new System.Drawing.Size(32, 24);
			this.btQuitar.TabIndex = 19;
			this.btQuitar.Text = "<-";
			this.btQuitar.Click += new System.EventHandler(this.btQuitar_Click);
			// 
			// btCambiaER
			// 
			this.btCambiaER.Location = new System.Drawing.Point(312, 9);
			this.btCambiaER.Name = "btCambiaER";
			this.btCambiaER.Size = new System.Drawing.Size(32, 24);
			this.btCambiaER.TabIndex = 20;
			this.btCambiaER.Text = "<>";
			this.btCambiaER.Click += new System.EventHandler(this.btCambiaER_Click);
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
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsVendedor_BeforeFill);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Vendedor = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.cdvPersonal;
			this.cmbVendedor.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn187.Header.VisiblePosition = 21;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 59;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn188.Width = 16;
			ultraGridColumn189.Header.VisiblePosition = 19;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn190.Header.VisiblePosition = 20;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn191.Header.VisiblePosition = 18;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn192.Header.VisiblePosition = 22;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn193.Header.VisiblePosition = 36;
			ultraGridColumn193.Hidden = true;
			ultraGridColumn194.Header.VisiblePosition = 25;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 0;
			ultraGridColumn195.Width = 38;
			ultraGridColumn196.Header.VisiblePosition = 1;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 35;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn198.Header.VisiblePosition = 24;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 13;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 32;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 7;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.Header.VisiblePosition = 6;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 3;
			ultraGridColumn203.Hidden = true;
			ultraGridColumn204.Header.VisiblePosition = 29;
			ultraGridColumn204.Hidden = true;
			ultraGridColumn205.Header.VisiblePosition = 28;
			ultraGridColumn205.Hidden = true;
			ultraGridColumn206.Header.VisiblePosition = 23;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn207.Header.VisiblePosition = 31;
			ultraGridColumn207.Hidden = true;
			ultraGridColumn208.Header.VisiblePosition = 10;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn209.Header.VisiblePosition = 33;
			ultraGridColumn209.Hidden = true;
			ultraGridColumn210.Header.VisiblePosition = 14;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn211.Header.VisiblePosition = 15;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn212.Header.VisiblePosition = 17;
			ultraGridColumn212.Hidden = true;
			ultraGridColumn213.Header.VisiblePosition = 26;
			ultraGridColumn213.Hidden = true;
			ultraGridColumn214.Header.VisiblePosition = 5;
			ultraGridColumn214.Hidden = true;
			ultraGridColumn215.Header.VisiblePosition = 37;
			ultraGridColumn215.Hidden = true;
			ultraGridColumn216.Header.VisiblePosition = 38;
			ultraGridColumn216.Hidden = true;
			ultraGridColumn217.Header.VisiblePosition = 27;
			ultraGridColumn217.Hidden = true;
			ultraGridColumn218.Header.VisiblePosition = 30;
			ultraGridColumn218.Hidden = true;
			ultraGridColumn219.Header.VisiblePosition = 12;
			ultraGridColumn219.Hidden = true;
			ultraGridColumn220.Header.VisiblePosition = 9;
			ultraGridColumn220.Hidden = true;
			ultraGridColumn221.Header.VisiblePosition = 11;
			ultraGridColumn221.Hidden = true;
			ultraGridColumn222.Header.VisiblePosition = 16;
			ultraGridColumn222.Hidden = true;
			ultraGridColumn223.Header.VisiblePosition = 2;
			ultraGridColumn223.Hidden = true;
			ultraGridColumn224.Header.VisiblePosition = 34;
			ultraGridColumn224.Hidden = true;
			ultraGridColumn225.Header.VisiblePosition = 4;
			ultraGridColumn225.Hidden = true;
			ultraGridColumn226.Header.VisiblePosition = 8;
			ultraGridColumn226.Hidden = true;
			ultraGridColumn227.Header.VisiblePosition = 39;
			ultraGridColumn227.Hidden = true;
			ultraGridColumn227.Width = 8;
			ultraGridColumn228.Header.VisiblePosition = 40;
			ultraGridColumn228.Hidden = true;
			ultraGridColumn228.Width = 8;
			ultraGridColumn229.Header.VisiblePosition = 41;
			ultraGridColumn229.Hidden = true;
			ultraGridColumn229.Width = 8;
			ultraGridColumn230.Header.VisiblePosition = 42;
			ultraGridColumn230.Hidden = true;
			ultraGridColumn230.Width = 8;
			ultraGridColumn231.Header.VisiblePosition = 43;
			ultraGridColumn231.Hidden = true;
			ultraGridColumn231.Width = 8;
			ultraGridColumn232.Header.VisiblePosition = 44;
			ultraGridColumn232.Hidden = true;
			ultraGridColumn232.Width = 8;
			ultraGridColumn233.Header.VisiblePosition = 45;
			ultraGridColumn233.Hidden = true;
			ultraGridColumn233.Width = 8;
			ultraGridColumn234.Header.VisiblePosition = 46;
			ultraGridColumn234.Hidden = true;
			ultraGridColumn234.Width = 8;
			ultraGridColumn235.Header.VisiblePosition = 47;
			ultraGridColumn235.Hidden = true;
			ultraGridColumn235.Width = 8;
			ultraGridColumn236.Header.VisiblePosition = 48;
			ultraGridColumn236.Hidden = true;
			ultraGridColumn236.Width = 8;
			ultraGridColumn237.Header.VisiblePosition = 49;
			ultraGridColumn237.Hidden = true;
			ultraGridColumn237.Width = 8;
			ultraGridColumn238.Header.VisiblePosition = 50;
			ultraGridColumn238.Hidden = true;
			ultraGridColumn238.Width = 8;
			ultraGridColumn239.Header.VisiblePosition = 51;
			ultraGridColumn239.Hidden = true;
			ultraGridColumn239.Width = 8;
			ultraGridColumn240.Header.VisiblePosition = 52;
			ultraGridColumn240.Hidden = true;
			ultraGridColumn240.Width = 8;
			ultraGridColumn241.Header.VisiblePosition = 53;
			ultraGridColumn241.Hidden = true;
			ultraGridColumn241.Width = 8;
			ultraGridColumn242.Header.VisiblePosition = 54;
			ultraGridColumn242.Hidden = true;
			ultraGridColumn242.Width = 8;
			ultraGridColumn243.Header.VisiblePosition = 55;
			ultraGridColumn243.Hidden = true;
			ultraGridColumn243.Width = 8;
			ultraGridColumn244.Header.VisiblePosition = 56;
			ultraGridColumn244.Hidden = true;
			ultraGridColumn244.Width = 14;
			ultraGridColumn245.Header.VisiblePosition = 57;
			ultraGridColumn245.Hidden = true;
			ultraGridColumn245.Width = 14;
			ultraGridColumn246.Header.VisiblePosition = 58;
			ultraGridColumn246.Hidden = true;
			ultraGridColumn246.Width = 17;
			ultraGridColumn247.Header.VisiblePosition = 60;
			ultraGridColumn247.Hidden = true;
			ultraGridColumn247.Width = 27;
			ultraGridColumn248.Header.VisiblePosition = 61;
			ultraGridColumn248.Hidden = true;
			ultraGridColumn248.Width = 60;
			ultraGridColumn249.Header.VisiblePosition = 62;
			ultraGridColumn249.Hidden = true;
			ultraGridColumn249.Width = 43;
			ultraGridColumn250.Header.VisiblePosition = 63;
			ultraGridColumn250.Hidden = true;
			ultraGridColumn250.Width = 30;
			ultraGridColumn251.Header.VisiblePosition = 64;
			ultraGridColumn251.Hidden = true;
			ultraGridColumn251.Width = 35;
			ultraGridColumn252.Header.VisiblePosition = 65;
			ultraGridColumn252.Hidden = true;
			ultraGridColumn252.Width = 43;
			ultraGridColumn253.Header.VisiblePosition = 66;
			ultraGridColumn253.Hidden = true;
			ultraGridColumn253.Width = 50;
			ultraGridColumn254.Header.VisiblePosition = 67;
			ultraGridColumn254.Hidden = true;
			ultraGridColumn254.Width = 8;
			ultraGridColumn255.Header.VisiblePosition = 68;
			ultraGridColumn255.Hidden = true;
			ultraGridColumn255.Width = 8;
			ultraGridColumn256.Header.VisiblePosition = 69;
			ultraGridColumn256.Hidden = true;
			ultraGridColumn256.Width = 8;
			ultraGridColumn257.Header.VisiblePosition = 70;
			ultraGridColumn257.Hidden = true;
			ultraGridColumn257.Width = 8;
			ultraGridColumn258.Header.VisiblePosition = 71;
			ultraGridColumn258.Hidden = true;
			ultraGridColumn258.Width = 8;
			ultraGridColumn259.Header.VisiblePosition = 72;
			ultraGridColumn259.Hidden = true;
			ultraGridColumn259.Width = 8;
			ultraGridColumn260.Header.VisiblePosition = 73;
			ultraGridColumn260.Hidden = true;
			ultraGridColumn260.Width = 8;
			ultraGridColumn261.Header.VisiblePosition = 74;
			ultraGridColumn261.Hidden = true;
			ultraGridColumn261.Width = 8;
			ultraGridColumn262.Header.VisiblePosition = 75;
			ultraGridColumn262.Hidden = true;
			ultraGridColumn262.Width = 8;
			ultraGridColumn263.Header.VisiblePosition = 76;
			ultraGridColumn263.Hidden = true;
			ultraGridColumn263.Width = 8;
			ultraGridColumn264.Header.VisiblePosition = 77;
			ultraGridColumn264.Hidden = true;
			ultraGridColumn264.Width = 8;
			ultraGridColumn265.Header.VisiblePosition = 78;
			ultraGridColumn265.Hidden = true;
			ultraGridColumn265.Width = 8;
			ultraGridColumn266.Header.VisiblePosition = 79;
			ultraGridColumn266.Hidden = true;
			ultraGridColumn266.Width = 8;
			ultraGridColumn267.Header.VisiblePosition = 80;
			ultraGridColumn267.Hidden = true;
			ultraGridColumn267.Width = 8;
			ultraGridColumn268.Header.VisiblePosition = 81;
			ultraGridColumn268.Hidden = true;
			ultraGridColumn268.Width = 8;
			ultraGridColumn269.Header.VisiblePosition = 82;
			ultraGridColumn269.Hidden = true;
			ultraGridColumn269.Width = 8;
			ultraGridColumn270.Header.VisiblePosition = 83;
			ultraGridColumn270.Hidden = true;
			ultraGridColumn270.Width = 8;
			ultraGridColumn271.Header.VisiblePosition = 84;
			ultraGridColumn271.Hidden = true;
			ultraGridColumn271.Width = 8;
			ultraGridColumn272.Header.VisiblePosition = 85;
			ultraGridColumn272.Hidden = true;
			ultraGridColumn272.Width = 8;
			ultraGridColumn273.Header.VisiblePosition = 86;
			ultraGridColumn273.Hidden = true;
			ultraGridColumn273.Width = 8;
			ultraGridColumn274.Header.VisiblePosition = 87;
			ultraGridColumn274.Hidden = true;
			ultraGridColumn274.Width = 8;
			ultraGridColumn275.Header.VisiblePosition = 88;
			ultraGridColumn275.Hidden = true;
			ultraGridColumn275.Width = 8;
			ultraGridColumn276.Header.VisiblePosition = 89;
			ultraGridColumn276.Hidden = true;
			ultraGridColumn276.Width = 8;
			ultraGridColumn277.Header.VisiblePosition = 90;
			ultraGridColumn277.Hidden = true;
			ultraGridColumn277.Width = 8;
			ultraGridColumn278.Header.VisiblePosition = 91;
			ultraGridColumn278.Hidden = true;
			ultraGridColumn278.Width = 8;
			ultraGridColumn279.Header.VisiblePosition = 92;
			ultraGridColumn279.Hidden = true;
			ultraGridColumn279.Width = 8;
			ultraGridColumn280.Header.VisiblePosition = 93;
			ultraGridColumn280.Hidden = true;
			ultraGridColumn280.Width = 8;
			ultraGridColumn281.Header.VisiblePosition = 94;
			ultraGridColumn281.Hidden = true;
			ultraGridColumn281.Width = 16;
			ultraGridColumn282.Header.VisiblePosition = 95;
			ultraGridColumn282.Width = 17;
			ultraGridColumn283.Header.VisiblePosition = 96;
			ultraGridColumn283.Width = 11;
			ultraGridColumn284.Header.VisiblePosition = 97;
			ultraGridColumn284.Width = 11;
			ultraGridColumn285.Header.VisiblePosition = 98;
			ultraGridColumn285.Width = 11;
			ultraGridColumn286.Header.VisiblePosition = 99;
			ultraGridColumn286.Width = 23;
			ultraGridColumn287.Header.VisiblePosition = 100;
			ultraGridColumn287.Width = 21;
			ultraGridColumn288.Header.VisiblePosition = 101;
			ultraGridColumn288.Width = 21;
			ultraGridColumn289.Header.VisiblePosition = 102;
			ultraGridColumn289.Width = 19;
			ultraGridColumn290.Header.VisiblePosition = 103;
			ultraGridColumn290.Width = 16;
			ultraGridColumn291.Header.VisiblePosition = 104;
			ultraGridColumn291.Width = 17;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190,
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201,
																										 ultraGridColumn202,
																										 ultraGridColumn203,
																										 ultraGridColumn204,
																										 ultraGridColumn205,
																										 ultraGridColumn206,
																										 ultraGridColumn207,
																										 ultraGridColumn208,
																										 ultraGridColumn209,
																										 ultraGridColumn210,
																										 ultraGridColumn211,
																										 ultraGridColumn212,
																										 ultraGridColumn213,
																										 ultraGridColumn214,
																										 ultraGridColumn215,
																										 ultraGridColumn216,
																										 ultraGridColumn217,
																										 ultraGridColumn218,
																										 ultraGridColumn219,
																										 ultraGridColumn220,
																										 ultraGridColumn221,
																										 ultraGridColumn222,
																										 ultraGridColumn223,
																										 ultraGridColumn224,
																										 ultraGridColumn225,
																										 ultraGridColumn226,
																										 ultraGridColumn227,
																										 ultraGridColumn228,
																										 ultraGridColumn229,
																										 ultraGridColumn230,
																										 ultraGridColumn231,
																										 ultraGridColumn232,
																										 ultraGridColumn233,
																										 ultraGridColumn234,
																										 ultraGridColumn235,
																										 ultraGridColumn236,
																										 ultraGridColumn237,
																										 ultraGridColumn238,
																										 ultraGridColumn239,
																										 ultraGridColumn240,
																										 ultraGridColumn241,
																										 ultraGridColumn242,
																										 ultraGridColumn243,
																										 ultraGridColumn244,
																										 ultraGridColumn245,
																										 ultraGridColumn246,
																										 ultraGridColumn247,
																										 ultraGridColumn248,
																										 ultraGridColumn249,
																										 ultraGridColumn250,
																										 ultraGridColumn251,
																										 ultraGridColumn252,
																										 ultraGridColumn253,
																										 ultraGridColumn254,
																										 ultraGridColumn255,
																										 ultraGridColumn256,
																										 ultraGridColumn257,
																										 ultraGridColumn258,
																										 ultraGridColumn259,
																										 ultraGridColumn260,
																										 ultraGridColumn261,
																										 ultraGridColumn262,
																										 ultraGridColumn263,
																										 ultraGridColumn264,
																										 ultraGridColumn265,
																										 ultraGridColumn266,
																										 ultraGridColumn267,
																										 ultraGridColumn268,
																										 ultraGridColumn269,
																										 ultraGridColumn270,
																										 ultraGridColumn271,
																										 ultraGridColumn272,
																										 ultraGridColumn273,
																										 ultraGridColumn274,
																										 ultraGridColumn275,
																										 ultraGridColumn276,
																										 ultraGridColumn277,
																										 ultraGridColumn278,
																										 ultraGridColumn279,
																										 ultraGridColumn280,
																										 ultraGridColumn281,
																										 ultraGridColumn282,
																										 ultraGridColumn283,
																										 ultraGridColumn284,
																										 ultraGridColumn285,
																										 ultraGridColumn286,
																										 ultraGridColumn287,
																										 ultraGridColumn288,
																										 ultraGridColumn289,
																										 ultraGridColumn290,
																										 ultraGridColumn291});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(80, 317);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(224, 21);
			this.cmbVendedor.TabIndex = 21;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 317);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(52, 16);
			this.label4.TabIndex = 22;
			this.label4.Text = "Vendedor";
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btTotal.Location = new System.Drawing.Point(488, 317);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(40, 25);
			this.btTotal.TabIndex = 23;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(544, 317);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.Size = new System.Drawing.Size(100, 22);
			this.txtTotal.TabIndex = 24;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(512, 72);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(136, 25);
			this.ultraButton1.TabIndex = 25;
			this.ultraButton1.Text = "Transformacion Directa";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 72);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(47, 16);
			this.label5.TabIndex = 26;
			this.label5.Text = "Comprb.";
			// 
			// txtComprobante
			// 
			this.txtComprobante.Location = new System.Drawing.Point(64, 72);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(80, 22);
			this.txtComprobante.TabIndex = 27;
			// 
			// chkServicio
			// 
			this.chkServicio.Location = new System.Drawing.Point(216, 72);
			this.chkServicio.Name = "chkServicio";
			this.chkServicio.Size = new System.Drawing.Size(80, 21);
			this.chkServicio.TabIndex = 28;
			this.chkServicio.Text = "Servicio";
			this.chkServicio.Visible = false;
			// 
			// cmbProyecto
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance14;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn292.Header.VisiblePosition = 0;
			ultraGridColumn293.Header.VisiblePosition = 1;
			ultraGridColumn293.Width = 408;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn292,
																										 ultraGridColumn293});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(408, 40);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(176, 21);
			this.cmbProyecto.TabIndex = 88;
			this.cmbProyecto.ValueMember = "";
			this.cmbProyecto.Visible = false;
			// 
			// lblProyecto
			// 
			this.lblProyecto.AutoSize = true;
			this.lblProyecto.Location = new System.Drawing.Point(352, 40);
			this.lblProyecto.Name = "lblProyecto";
			this.lblProyecto.Size = new System.Drawing.Size(48, 16);
			this.lblProyecto.TabIndex = 89;
			this.lblProyecto.Text = "Centro C";
			this.lblProyecto.Visible = false;
			// 
			// lblCaja
			// 
			this.lblCaja.AutoSize = true;
			this.lblCaja.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCaja.ForeColor = System.Drawing.Color.Red;
			this.lblCaja.Location = new System.Drawing.Point(592, 8);
			this.lblCaja.Name = "lblCaja";
			this.lblCaja.Size = new System.Drawing.Size(39, 22);
			this.lblCaja.TabIndex = 90;
			this.lblCaja.Text = "Tipo";
			this.lblCaja.Visible = false;
			// 
			// btTSP
			// 
			this.btTSP.Location = new System.Drawing.Point(312, 312);
			this.btTSP.Name = "btTSP";
			this.btTSP.Size = new System.Drawing.Size(40, 25);
			this.btTSP.TabIndex = 91;
			this.btTSP.Text = "TSP";
			this.btTSP.Click += new System.EventHandler(this.btTSP_Click);
			// 
			// txtNombre
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance15;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(64, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(232, 22);
			this.txtNombre.TabIndex = 92;
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// cmbCliente
			// 
			this.cmbCliente.Location = new System.Drawing.Point(64, 40);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.Ascending),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idCliente", "", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.MaxLength = 400;
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.PopupWidth = 400;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(248, 20);
			this.cmbCliente.TabIndex = 142;
			this.cmbCliente.EditValueChanged += new System.EventHandler(this.cmbCliente_EditValueChanged);
			// 
			// TransferenciaArticulos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(664, 366);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btTSP);
			this.Controls.Add(this.lblCaja);
			this.Controls.Add(this.lblProyecto);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.chkServicio);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.btCambiaER);
			this.Controls.Add(this.btQuitar);
			this.Controls.Add(this.btAgregarTodo);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.btTransformar);
			this.Controls.Add(this.btQuitarTodo);
			this.Controls.Add(this.btAgregar);
			this.Controls.Add(this.chkActualizaPrecio);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.cmbTipoRecibe);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "TransferenciaArticulos";
			this.Text = "Comprobante";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TransferenciaArticulos_Closing);
			this.Load += new System.EventHandler(this.TransferenciaArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTmpArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompraNumero_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			this.cdsCompraNumero.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(cmbTipo, "");
			errorProvider.SetError(cmbTipoRecibe, "");
			errorProvider.SetError(ultraGrid2, "");	
			errorProvider.SetError(cmbProyecto, "");				
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (cmbCliente.EditValue == null)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				bOk = false;
			}
			if (cmbTipo.Value == null)
			{
				errorProvider.SetError(cmbTipo, "Ingrese Tipo");
				bOk = false;
			}
			if(Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'"))
			{
				if (cmbProyecto.Value == null)
				{
					errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
					bOk = false;
				}
			}
			return bOk;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			string stFiltro = "idTipoFactura = " + cmbTipo.Value.ToString().Trim();
			stFiltro += " And idCliente = " + this.cmbCliente.EditValue.ToString().Trim();
			stFiltro += " And Total > 0 And Borrar = 0 ";

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTTSC'")) 				
			{
			stFiltro += " And isnull(mensaje1,'') <> 'T' ";
			}
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'")) 				
			{
				stFiltro += " And compra.idProyecto = "+this.cmbProyecto.Value.ToString().Trim();
			}

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPCO'")) 				
			{
				stFiltro += " And ISNULL(p_IDORIGEN,0)  = 0 ";
			}

			if (txtComprobante.Text.Trim().Length > 0)
			{
				stFiltro += " And Comprobante = '" + txtComprobante.Text.Trim() + "' ";

			}

			C1.Data.FilterCondition fcPadre = new 
			C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			C1.Data.FilterCondition fcHija = new 
			C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
			fcTotal.Add(fcHija);

			try
			{
				this.cdsCompra.Fill(fcTotal, false);
				ultraGrid1.Select();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "No existe datos de este cliente");
			}
		}

			C1.Data.C1DataRow drSeteoF;
			private void TransferenciaArticulos_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Transf. Articulos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			drSeteo = this.cdsSeteo.TableViews["Seteo"].Rows[0];
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			Cursor = Cursors.Default;
			// Desabilta cuando en factura de venta no se puede cambiar precios
			if (!Funcion.Permiso("277", cdsCompra))
			{
				ultraGrid2.Enabled = false;
				chkActualizaPrecio.Enabled = false;
				btAgregar.Enabled = false;
				btQuitar.Enabled = false;
//				ultraGrid2.Layout.Bands[0].Columns["Precio"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;

			}
				cargarComboTipos();
	
				if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FTRPCO'")) 				
				{
						btTransformar.Width =0; 
				}
				#region Cambios PUBLLIJOB
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'")) 				
				{			
					cargaProyecto();
					cmbProyecto.Visible = true;
					lblProyecto.Visible = true;	
					lblCaja.Visible= true;
					string stSelect="";							
					///carga combo cmbTipo Recibe
					stSelect = "select idtipofactura, Nombre from compranumero where idtipofactura in (1,19) order by idtipofactura ";
					cmbTipoRecibe.DataSource  = Funcion.dvProcedimiento(cdsArticulo,stSelect);
					cmbTipoRecibe.DisplayMember = "Nombre";
					cmbTipoRecibe.ValueMember = "idtipofactura";
					cmbTipoRecibe.Rows.Band.Columns["idtipofactura"].Hidden = true;

					
					///carga combo cmbTipo
					stSelect = "select idtipofactura, Nombre from compranumero where idtipofactura = (13) order by idtipofactura";
					cmbTipo.DataSource  = Funcion.dvProcedimiento(cdsArticulo,stSelect);
					cmbTipo.DisplayMember = "Nombre";
					cmbTipo.ValueMember = "idtipofactura";
					cmbTipo.Rows.Band.Columns["idtipofactura"].Hidden = true;
				}	
				#endregion Cambios PUBLLIJOB
				cmbTipoRecibe.Value = 1;
				cmbTipo.Value = 8;
				if (drSeteoF["VerNombreLista"] != DBNull.Value && (bool) drSeteoF["VerNombreLista"])
				{
					string stExec = "Exec ClienteCargaNombre 0";
					cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
					txtNombre.Width = 0;
				}
				else
					cmbCliente.Width = 0;

			}
		private void cargaProyecto()
		{
			string stSelect = "select idProyecto,Nombre,CentroCosto from Proyecto where Visible =1";
			cmbProyecto.DataSource  = Funcion.dvProcedimiento(cdsArticulo,stSelect);
			cmbProyecto.DisplayMember = "Nombre";
			cmbProyecto.ValueMember = "idProyecto";
			cmbProyecto.Rows.Band.Columns["idProyecto"].Hidden = true;
		}

		private void cargarComboTipos()
		{
			string stSelect="";
			///carga combo cmbTipo
			stSelect = "select idtipofactura, Nombre from compranumero order by idtipofactura";
			cmbTipo.DataSource  = Funcion.dvProcedimiento(cdsArticulo,stSelect);
			cmbTipo.DisplayMember = "Nombre";
			cmbTipo.ValueMember = "idtipofactura";
			cmbTipo.Rows.Band.Columns["idtipofactura"].Hidden = true;

			///carga combo cmbTipo Recibe
			stSelect = "select idtipofactura, Nombre from compranumero order by idtipofactura";
			cmbTipoRecibe.DataSource  = Funcion.dvProcedimiento(cdsArticulo,stSelect);
			cmbTipoRecibe.DisplayMember = "Nombre";
			cmbTipoRecibe.ValueMember = "idtipofactura";
			cmbTipoRecibe.Rows.Band.Columns["idtipofactura"].Hidden = true;
		}


		private bool VerificaAgregar(int idCompra)
		{
			// Verificacion de existencia de factura en la lista de la derecha
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ((int) dr.Cells["idCompra"].Value == idCompra)
				{
//					MessageBox.Show("Transaccion ya esta pasada", "Información");
					return false;
				}
			}
			return true;
		}

		private void btAgregar_Click(object sender, System.EventArgs e)
		{
			int IdCompra = 0;
			try
			{
				IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			}
			catch
			{
				errorProvider.SetError(ultraGrid1, "Escoja Transaccion para pasar el detalle");
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				IdCompra = (int) dr.Cells["idCompra"].Value;
				if (!VerificaAgregar(IdCompra)) return;

				string stSelect = "SELECT idArticulo, Cantidad, Precio, idDetCompra, " +
					" Impuesto, DescuentoArt, Notas " +
					"FROM DetCompra WHERE DetCompra.idCompra = " + IdCompra.ToString()
					+ " And Cantidad > 0 Order by idDetCompra";

				this.cdsCompra.Schema.Connections[0].Open();
				SqlDataReader miReader;
				C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
				SqlCommand cmdCant = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				miReader = cmdCant.ExecuteReader();
				try
				{
					while (miReader.Read())
					{
						ultraGrid2.DisplayLayout.Bands[0].AddNew();
						int IdArt = miReader.GetInt32(0);
						ultraGrid2.ActiveRow.Cells["idArticulo"].Value = IdArt;
						ultraGrid2.ActiveRow.Cells["Codigo"].Value = IdArt;
						ultraGrid2.ActiveRow.Cells["idCompra"].Value = IdCompra;
						ultraGrid2.ActiveRow.Cells["Cantidad"].Value = miReader.GetDouble(1);
						ultraGrid2.ActiveRow.Cells["Precio"].Value = miReader.GetDouble(2);
						ultraGrid2.ActiveRow.Cells["idDetCompra"].Value = miReader.GetInt32(3);
						ultraGrid2.ActiveRow.Cells["Impuesto"].Value = miReader.GetDouble(4);
						ultraGrid2.ActiveRow.Cells["Descuento"].Value = miReader.GetDouble(5);
						if (miReader.GetValue(6) != DBNull.Value)
							ultraGrid2.ActiveRow.Cells["Notas"].Value = miReader.GetString(6);
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Pasar Lineas");
				}
				miReader.Close();
			}
			this.cdsCompra.Schema.Connections[0].Close();
		}

		private void btQuitar_Click(object sender, System.EventArgs e)
		{
			ultraGrid2.ActiveRow.Delete(false);
		}

		C1.Data.C1DataRow drSeteo;
		int IdCompraCreada = 0;
		private void CreaCompra()
		{
			C1.Data.C1DataRow drCompra = cdsCompra.TableViews["Compra"].AddNew();
			if ((int) cmbTipoRecibe.Value == 1) // Factura de venta
				drCompra["idComprobante"] = 15;
			else
				drCompra["idComprobante"] = 1;

			drCompra["idCliente"] = (int) cmbCliente.EditValue;
			bool bCredito = Funcion.bEscalarSQL(cdsArticulo, "Select IsNull(VentaCredito, 0) From SeteoF", true);
			int iContadoCredito = 1;
			if (bCredito) iContadoCredito = 2;
			drCompra["ContadoCredito"] = iContadoCredito;
			drCompra["idFormaPago"] = 1;
			drCompra["Fecha"] = DateTime.Today;
			drCompra["idTipoFactura"] = cmbTipoRecibe.Value;
			drCompra["idSucursal"] = (int) this.drSeteo["NumEmpresa"];
			if (chkActualizaPrecio.Enabled == false || cmbVendedor.Value == null)
			{
				string stBuscaV = "Select idVendedor From Compra where idCompra = "
					+ ultraGrid2.Rows[0].Cells["idCompra"].Value.ToString();
				drCompra["idVendedor"] = Funcion.iEscalarSQL(cdsCompra, stBuscaV, false);
			}
			else
			{
				drCompra["idVendedor"] = (int) cmbVendedor.Value;
			}
			if (!Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRNBG'"))
//				if ((int)cmbTipoRecibe.Value == 1)
				{
					string stNumero = Funcion.NumeraFactura(cdsCompra, (int) cmbTipoRecibe.Value);
					if (stNumero.Trim().Length > 0) drCompra["Numero"] = stNumero;
				}

			#region cambios publijob crear factura como ntv
			// cuando escoja en el listado nota de venta se creara la factura con comprobante notaventa
			// y asignara el centro de costo que selecciono en el combo
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'")) 				
			{	
				if ((int) cmbTipoRecibe.Value == 19) // Nota de venta
				{
					drCompra["idTipoFactura"] = 1;
					drCompra["idComprobante"] = 2;
				}
				drCompra["idProyecto"] = cmbProyecto.Value; // Centro de costo

				//numeracion
				if ((int)cmbTipoRecibe.Value == 19 )
				{
					string stNumero = Funcion.NumeraFactura(cdsCompra, (int) cmbTipoRecibe.Value);
					if (stNumero.Trim().Length > 0) drCompra["Numero"] = stNumero;
				}
			}
			#endregion cambios publijob crear factura como ntv 
			cdsCompra.Update();
			IdCompraCreada = (int) drCompra["idCompra"];
		}


		private int iSignoDestino = 1;
		private bool bSignoDestino = true;
		private bool VerificaT()
		{
			bool bOk = true;
			BorraErrores();
			int IdTipoOrigen = int.Parse(cmbTipo.Value.ToString());
			int IdTipoDestino = int.Parse(cmbTipoRecibe.Value.ToString());
			string stSelect = "Select Signo From CompraNumero "
				+ " Where idTipoFactura = " + IdTipoOrigen.ToString();
			int iSigno1 = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			stSelect = string.Format("Exec TransferenciaArtDisminuye {0}, {1}", IdTipoOrigen, IdTipoDestino);

			if (IdTipoOrigen == 14) iSigno1 = 1;

			bSignoDestino = Funcion.bEscalarSQL(cdsCompra, stSelect, true);
			if (this.ultraGrid2.Rows.Count == 0)
			{
				errorProvider.SetError(ultraGrid2, "No Existen Articulos a Transformar");
				bOk = false;
			}

			if (bSignoDestino)
			{
				bCotizacion  = true;
			}
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTTSC'")) 				
			{
				bCotizacion = true;
			}
			// Cambio para Pichincha Directo
//			else
//			{
//				bCotizacion  = false;
//				if (iSigno1 != iSignoDestino)
//				{
//					errorProvider.SetError(cmbTipo, "No puede Transformar una "
//						+ cmbTipo.Text + " en " + cmbTipoRecibe.Text);
//					bOk = false;
//				}
//			}
			return bOk;
		}

		private bool VerificaTodo()
		{
			bool bOk = true;
			BorraErrores();
			int IdTipoOrigen = int.Parse(cmbTipo.Value.ToString());
			int IdTipoDestino = int.Parse(cmbTipoRecibe.Value.ToString());
			if (cmbCliente.EditValue == null)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				bOk = false;
			}
			string stSelect = "Select Signo From CompraNumero "
				+ " Where idTipoFactura = " + IdTipoOrigen.ToString();
			int iSigno1 = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			stSelect = "Select Signo From CompraNumero "
				+ " Where idTipoFactura = " + IdTipoDestino.ToString();

			if (IdTipoOrigen == 14) iSigno1 = 1;

			iSignoDestino = Funcion.iEscalarSQL(cdsCompra, stSelect, true);

				if (iSigno1 != iSignoDestino)
				{
					errorProvider.SetError(cmbTipo, "No puede Transformar una "
						+ cmbTipo.Text + " en " + cmbTipoRecibe.Text);
					bOk = false;
				}
			return bOk;
		}

		bool bCotizacion = false;
		
		CultureInfo us = new CultureInfo("en-US");
		private void btTransformar_Click(object sender, System.EventArgs e)
		{
			if (!VerificaT()) return;
			string stMensaje = "¿Desea transferir Artículos desde "
				+ cmbTipo.Text + " hasta " + cmbTipoRecibe.Text +"?";
			if (DialogResult.No == MessageBox.Show(stMensaje, "Información",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			CreaCompra();
			int idInicio = IdCompraCreada;
			int iNumLineas = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (iNumLineas >= (int) this.drSeteoF["NumeroLineas"])
				{
					CreaCompra();
					iNumLineas = 0;
				}
				iNumLineas ++;
				this.cdsCompra.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
				C1.Data.C1DataRow drDC = cdsCompra.TableViews["DetCompra"].AddNew();
				double dCantidad = (double) dr.Cells["Cantidad"].Value;
				double dPrecio = (double) dr.Cells["Precio"].Value;
				int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTTVS'"))
				{
					int IdArti = (int)dr.Cells["idArticulo"].Value;
					string stSelect = "Select saldo From Articulo "
						+ " Where idArticulo = " + IdArti.ToString();
					double dSaldo = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
					if (dCantidad > dSaldo)
					{
						dCantidad = dSaldo;
						MessageBox.Show("Cambio cantidad a saldo","Información" );
					}
				}	
				drDC["Cantidad"] = dCantidad;
				drDC["Precio"] = dPrecio;
				drDC["Signo"] = iSignoDestino;
				drDC["idArticulo"] = (int) dr.Cells["idArticulo"].Value;
				double dImpuesto = Funcion.dEscalarSQL(cdsArticulo, "Select Iva From Articulo Where idArticulo = " + dr.Cells["idArticulo"].Value);
////				drDC["idProyecto"] = (int) dr.Cells["idProyecto"].Value;
////				drDC["idSubProyecto"] = (int) dr.Cells["idSubProyecto"].Value;
				drDC["idCompra"] = IdCompraCreada;
//				drDC["Articulo"] = dr.Cells["Articulo"].Value.ToString();
//				drDC["Codigo"] = dr.Cells["Codigo"].Value.ToString();
				drDC["Impuesto"] = dImpuesto;
				drDC["DescuentoArt"] = (double) dr.Cells["Descuento"].Value;
				drDC["Notas"] = dr.Cells["Notas"].Value.ToString();
				
				if (!bCotizacion) // No es cotizacion Mantiene los Valores / o cuando nos es FTTSC
				{
					string stSelect = "UPDATE DetCompra SET Cantidad = " +
						" Cantidad - " + dCantidad.ToString() +
						", Pedido = Pedido + " + dCantidad.ToString() +
						" Where DetCompra.idDetCompra = " + IdDetCompra.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
				}
				#region Campos de Presupuesto de compra a trasformar
				if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FTRPCO'")) 				
				{	
					string stProced = "";
					string retorna ="";
					int retornaId = 0;
					//asigna la obra del IGB a la nueva Compra
					stProced = "select isnull(PresupuestoC,'') from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retorna = Funcion.sEscalarSQL(cdsArticulo,stProced);					 
					drDC["PresupuestoC"] = retorna;
					stProced = "select isnull(idPresupuestoC,0) from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retornaId = Funcion.iEscalarSQL(cdsArticulo,stProced);					 
					drDC["idPresupuestoC"] = retornaId;

					//asigna el capitulo del IGB a la nueva Compra
					stProced = "select isnull(Capitulo,'') from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retorna = Funcion.sEscalarSQL(cdsArticulo,stProced);					 
					drDC["Capitulo"] = retorna;					
					stProced = "select isnull(idCapitulo,0) from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retornaId = Funcion.iEscalarSQL(cdsArticulo,stProced);					 
					drDC["idCapitulo"] = retornaId;

					//asigna del subCapitulo del IGB a la nueva Compra
					stProced = "select isnull(SubCapitulo,'') from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retorna = Funcion.sEscalarSQL(cdsArticulo,stProced);					 
					drDC["SubCapitulo"] = retorna;
					stProced = "select isnull(idSubCapitulo,0) from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retornaId = Funcion.iEscalarSQL(cdsArticulo,stProced);					 
					drDC["idSubCapitulo"] = retornaId;

					//asigna del Rubro del IGB a la nueva Compra
					stProced = "select isnull(PresupuestoCDet,'') from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retorna = Funcion.sEscalarSQL(cdsArticulo,stProced);					 
					drDC["PresupuestoCDet"] = retorna;
					stProced = "select isnull(idPresupuestoCDet,'') from detcompra where idDetcompra = "+IdDetCompra.ToString();
					retornaId = Funcion.iEscalarSQL(cdsArticulo,stProced);					 
					drDC["idPresupuestoCDet"] = retornaId;

					string stUpdate = "Update DetCompra set Signo = 0 where isnull(idPresupuestoC,0)= 0 ";
					Funcion.EjecutaSQL(cdsCompra,stUpdate);

					//ingresa el valor de la cantidad a pasar en  el campo pedido
					string stSelect = "UPDATE DetCompra SET Pedido = " + dCantidad.ToString() +
						" Where DetCompra.idDetCompra = " + IdDetCompra.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			
					// pone el impuesto que va desde el origen no toma el del articulo
					stProced = "select isnull(impuesto,0) from detcompra where idDetcompra = "+IdDetCompra.ToString();
					double imp = Funcion.dEscalarSQL(cdsArticulo,stProced);					 
					drDC["Impuesto"] = imp;

				}
				#endregion Campos de Presupuesto de compra a trasformar
				this.cdsCompra.Schema.Connections[0].Close();
			}
			cdsCompra.Update();


			// Grabacion de Totales de Facturas Nuevas
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion1 = cdsCompra.Schema.Connections[0];
			for (int i = idInicio; i<=IdCompraCreada; i++)
			{
	
				string stSelect1 = "EXEC FacturaTotal " + i.ToString();
				SqlCommand cmdCant1 = new SqlCommand (
					stSelect1, (SqlConnection) miConeccion1.DbConnection);
				try
				{
					cmdCant1.ExecuteNonQuery();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Grabar Compra");
				}
			}
			this.cdsCompra.Schema.Connections[0].Close();

			// Grabacion de Totales de Trasnsferencias Enviadas
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion1 = cdsCompra.Schema.Connections[0];
			int idCompraAntes = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				int idCompraActual = (int) dr.Cells["idCompra"].Value;
				if (idCompraAntes == idCompraActual) continue;
				idCompraAntes = idCompraActual;

				if (chkServicio.Checked)
				{
					string stUpdate = "Update Compra Set Servicio = (SubTotalIva + SubTotalExcento) "
						+ "Where idCompra = " + idCompraActual.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stUpdate);
				}

				string stSelect1 = "EXEC FacturaTotal " + idCompraActual.ToString();
				SqlCommand cmdCant1 = new SqlCommand (
					stSelect1, (SqlConnection) miConeccion1.DbConnection);
				try
				{
					cmdCant1.ExecuteNonQuery();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Grabar Compra origen");
				}
			}
			this.cdsCompra.Schema.Connections[0].Close();

			btQuitarTodo_Click(this, e);
			btFiltro_Click(this, e);
		#region PubliJob Actualiza lista de articulos al trasferir
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FCVCC'")) 				
			{				
				string stSelect = "Update DetCompra Set DetCompra.Codigo = Articulo.Codigo "
					+ ", DetCompra.Articulo = Articulo.Articulo "
					+ "From DetCompra Inner Join Articulo On DetCompra.idArticulo = Articulo.idArticulo "
					+ "Where IsNull(DetCompra.Codigo, '') = '' Or IsNull(DetCompra.Articulo, '') = ''";
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			}
			#endregion PubliJob Actualiza lista de articulos al trasferir
			MessageBox.Show(cmbTipoRecibe.Text + " Creada", "Información");
		}

		private void btQuitarTodo_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				dr.Delete(false);
			}		
		}

		private void btAgregarTodo_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
        "¿Desea copiar el contenido de todas las transacciones?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question)) return;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Activate();
				btAgregar_Click(this, e);
			}
		}

		private void ultraGrid1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				int IdCompra = 0;
				try
				{
					IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
				}
				catch
				{
					MessageBox.Show("Seleccione la Transaccion a Mostrar", "Información");
					return;
				}
				Cursor = Cursors.WaitCursor;
//				Compra miCompra = new Compra((int) cmbTipo.Value, IdCompra);
//				miCompra.MdiParent = this.MdiParent;
//				miCompra.Show();
				Cursor = Cursors.Default;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
		}

		private void cmbCliente_ValueChanged(object sender, System.EventArgs e)
		{
			btQuitarTodo_Click(this, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsVendedor_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			double dExcento = 0;
			double dIva = 0;
			double dDesc0 = 0;
			double dDesc = 0;
			double dImpuesto = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ((double) dr.Cells["Impuesto"].Value == 0)
				{
					dExcento += (double) dr.Cells["Precio"].Value
						* (double) dr.Cells["Cantidad"].Value;
					dDesc0 += (double) dr.Cells["Descuento"].Value;
				}
				else
				{
					dIva += (double) dr.Cells["Precio"].Value
						* (double) dr.Cells["Cantidad"].Value;
					dDesc += (double) dr.Cells["Descuento"].Value;
					dImpuesto += (double) dr.Cells["Impuesto"].Value * (double) dr.Cells["Precio"].Value
						* (double) dr.Cells["Cantidad"].Value / 100.0;
				}
			}
			txtTotal.Value = dExcento + dIva - dDesc - dDesc0 + dImpuesto;
		}

		private void TransferenciaArticulos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Transf. Articulos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}


		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

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
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}


		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
			{
				if (ultraGrid2.ActiveCell == null) return;
				try
				{
					if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
					switch(e.KeyValue)
					{
						case (int) Keys.PageDown:
							if(ultraGrid2.ActiveCell.DroppedDown == false)
								ultraGrid2.ActiveCell.DroppedDown = true;
							break;

						case (int) Keys.Up:
							if (ultraGrid2.ActiveCell.DroppedDown) return;

							ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							if (ultraGrid2.ActiveCell.DroppedDown) return;
			
							ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Enter:
							ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(ultraGrid2.ActiveCell.DroppedDown == false)
								ultraGrid2.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
					}
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message);
					//				MessageBox.Show(ex.Message, "Tecla");
				}

			}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (!VerificaT()) return;

			string stMensaje = "¿Desea transferir todos los Artículos desde "
				+ cmbTipo.Text + " hasta " + cmbTipoRecibe.Text +"?";
			
			if (DialogResult.No == MessageBox.Show(stMensaje, 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			stMensaje = "r";
			int iTipo = (int) cmbTipo.Value;
			int iTipoRecibe = (int) cmbTipoRecibe.Value;
			int iCliente = (int) cmbCliente.EditValue;
			string stExec = "Exec FacturaTransforma " + iTipo.ToString() +
				", " + iCliente + ", " + iTipoRecibe.ToString();
			string stMensajeError = Funcion.sEscalarSQL(cdsArticulo, stExec, true);
			if (stMensajeError.Length > 0)
			{
				MessageBox.Show(stMensajeError, "Error");
			}
			else
			{
				MessageBox.Show("Transformacion Generada", "Información");
			}
			btQuitarTodo_Click(this, e);
			btFiltro_Click(this, e);
		}

		private void btCambiaER_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btFiltro_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
	if (e.Button != MouseButtons.Right) return;
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCOTRA'")) 				
			{
				string stFiltro = "idTipoFactura = " + cmbTipo.Value.ToString().Trim();
				stFiltro += " And Total > 0 And Borrar = 0 ";
				
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'")) 				
				{
					errorProvider.SetError(cmbProyecto, "");
					if (cmbProyecto.Value == null)
					{
						errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
					return;
					}

					
					stFiltro += " And compra.idProyecto = "+this.cmbProyecto.Value.ToString().Trim();
				}
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);
				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
				fcTotal.Add(fcHija);
				try
				{
					this.cdsCompra.Fill(fcTotal, false);
					ultraGrid1.Select();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "No existe datos");
				}
			
			}
		
		}

		private void cmbTipoRecibe_TextChanged(object sender, System.EventArgs e)
		{
		if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVCC'")) 				
			{
				lblCaja.Visible = true;
				int iii= int.Parse(cmbTipoRecibe.Value.ToString());
				if(iii == 1)
					lblCaja.Text = "a Caja 1";
				else
					lblCaja.Text = "a Caja 2";			
			}
		}

		private void btTSP_Click(object sender, System.EventArgs e)
		{
			if (!VerificaT()) return;
			string stMensaje = "¿Desea transferir Artículos desde "
				+ cmbTipo.Text + " hasta " + cmbTipoRecibe.Text +"?";
			if (DialogResult.No == MessageBox.Show(stMensaje, "Información",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			CreaCompra();
			int idInicio = IdCompraCreada;
			int iNumLineas = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (iNumLineas >= (int) this.drSeteoF["NumeroLineas"])
				{
					CreaCompra();
					iNumLineas = 0;
				}
				iNumLineas ++;

				double dCantidad = (double) dr.Cells["Cantidad"].Value;
				double dPrecio = (double) dr.Cells["Precio"].Value;
				int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;
				double dDescuento = (double) dr.Cells["Descuento"].Value;
				string stNotas = dr.Cells["Notas"].Value.ToString();
				string stExec = String.Format("Exec ArticuloTransferencia {0}, {1}, {2}, {3}, {4}, '{5}' ",
					IdDetCompra, dCantidad.ToString("0.00000", us), dPrecio.ToString("0.00000", us), IdCompraCreada, 
					dDescuento.ToString("0.00000", us), stNotas);
				stMensaje = Funcion.sEscalarSQL(cdsArticulo, stExec);
				if (stMensaje.Length > 0)
					MessageBox.Show(stMensaje, "Informacion");
			}

			// Grabacion de Totales de Trasnsferencias Enviadas
			int idCompraAntes = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				int idCompraActual = (int) dr.Cells["idCompra"].Value;
				if (idCompraAntes == idCompraActual) continue;
				idCompraAntes = idCompraActual;

				string stSelect1 = "EXEC FacturaTotal " + idCompraActual.ToString();
				Funcion.EjecutaSQL(cdsArticulo, stSelect1);
			}

			btQuitarTodo_Click(this, e);
			btFiltro_Click(this, e);
			MessageBox.Show(cmbTipoRecibe.Text + " Creada", "Información");

		  string stExec1 = String.Format("Exec ArticuloTransferenciaFinal {0}", IdCompraCreada );
			stMensaje = Funcion.sEscalarSQL(cdsArticulo, stExec1);
			if (stMensaje.Length > 0)
				MessageBox.Show(stMensaje, "Informacion");

		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stFiltro = "0";
			int IdTipoFactura = (int) cmbTipo.Value;
			if (Funcion.EsProveedor(IdTipoFactura))
				stFiltro = "1";

			string stNombre = txtNombre.Text.Trim();
			if (stNombre.Length == 0) return;
			string stSelect = string.Format("Exec ClienteListaCuenta '{0}', {1}, 0", stNombre, stFiltro);
			int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
			int IdCliente = 0;
			if (iCuenta >= 0)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(stNombre, IdTipoFactura, 0, 0, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						IdCliente = miBusqueda.idCliente;
						txtNombre.Text = miBusqueda.Nombre;
						cmbCliente.EditValue = IdCliente;
					}
					else 
						e.Cancel=true;
				}
			}
			else // envia en negativo el idCliente cuando existe solo 1
			{
				IdCliente = -iCuenta;
				stSelect = "Select Nombre From Cliente Where idCliente = " + IdCliente.ToString();
				cmbCliente.EditValue = IdCliente;
			}

		}

		private void cmbCliente_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
